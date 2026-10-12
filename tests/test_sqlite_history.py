import asyncio
import csv
import json
import os
import tempfile
import unittest
from io import BytesIO, StringIO
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from backend import chat_history
from backend.routers import (
    ChatRequest,
    build_prompt,
    build_spreadsheet_answer,
    build_tool_inventory_list,
    build_tool_inventory_summary,
    chat,
    chat_stream,
    classify_chat_intent,
    download,
    download_maximo_object,
    _is_downloadable_document_url,
    find_maximo_document_download,
    get_maximo_chat_answer,
    get_maximo_summary_answer,
    maximo_asset_number,
    maximo_document_download_requested,
    maximo_collection_query,
    maximo_collection_count_requested,
    maximo_collection_where,
    maximo_download_tool_requested,
    maximo_download_object_requested,
    maximo_select_fields,
    maximo_summary_target,
    resolve_maximo_download_request,
    maximo_object_query,
    maximo_reference,
    tool_inventory_summary_requested,
    is_general_question,
    maximo_work_order_number
)
from backend.services import rag_service
from backend.services.file_service import (
    build_file_context,
    _normalized_tabular_records,
    build_spreadsheet_asset_count_answer,
    build_spreadsheet_asset_filter_count_answer,
    build_spreadsheet_count_answer,
    build_spreadsheet_column_count_answer,
    build_spreadsheet_distinct_column_count_answer,
    build_spreadsheet_distinct_location_count_answer,
    build_spreadsheet_location_count_answer,
    build_spreadsheet_location_answer,
    build_spreadsheet_item_count_answer,
    build_spreadsheet_priority_answer,
    build_spreadsheet_relationship_list_answer,
    build_spreadsheet_status_answer
)
from backend.services.file_service import load_file


class SqliteHistoryStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.original_chat_dir = chat_history.CHAT_DIR
        self.original_knowledge_db_path = rag_service.KNOWLEDGE_DB_PATH
        chat_history.CHAT_DIR = self.temp_dir
        chat_history.DB_PATH = os.path.join(self.temp_dir, "chat_history.db")
        rag_service.KNOWLEDGE_DB_PATH = os.path.join(self.temp_dir, "knowledge.sqlite3")
        chat_history.init_db()
        rag_service.init_knowledge_db()

    def tearDown(self):
        test_db_path = chat_history.DB_PATH
        chat_history.CHAT_DIR = self.original_chat_dir
        chat_history.DB_PATH = os.path.join(chat_history.CHAT_DIR, "chat_history.db")
        rag_service.KNOWLEDGE_DB_PATH = self.original_knowledge_db_path
        if os.path.exists(test_db_path):
            os.remove(test_db_path)

    def test_save_and_load_chat_round_trip(self):
        title = "hello world"
        messages = [{"user": "hi", "assistant": "hello"}]

        filename = chat_history.save_chat(title, messages)

        self.assertTrue(filename.endswith(".json"))
        self.assertIn("hello", filename)
        self.assertEqual(chat_history.list_history(), [filename])
        self.assertEqual(chat_history.load_chat(filename), messages)

    def test_load_file_extracts_excel_text(self):
        filepath = os.path.join(self.temp_dir, "sample.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Item", "Description"])
        sheet.append(["ITEM-1", "Filter"])
        workbook.save(filepath)

        text = load_file(filepath)

        self.assertIn("Item | Description", text)
        self.assertIn("ITEM-1 | Filter", text)

    def test_tabular_normalization_expands_y_n_boolean_values(self):
        filepath = os.path.join(self.temp_dir, "tools.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Tool", "Rotating", "Is M&TE"])
        sheet.append(["TOOL-1", "Y", "N"])
        sheet.append(["TOOL-2", "N", "N"])
        workbook.save(filepath)

        documents, metadata = _normalized_tabular_records(filepath)

        self.assertIn("Rotating=Yes", documents[0])
        self.assertIn("Is M&TE=No", documents[0])
        self.assertEqual(metadata[0]["attr_rotating"], "Yes")
        self.assertEqual(metadata[1]["attr_rotating"], "No")
        self.assertEqual(metadata[0]["attr_is_m_te"], "No")

    def test_delete_chat_removes_row(self):
        filename = chat_history.save_chat("delete me", [{"user": "a", "assistant": "b"}])

        self.assertTrue(chat_history.delete_chat(filename))
        self.assertEqual(chat_history.list_history(), [])
        self.assertEqual(chat_history.load_chat(filename), [])

    def test_title_with_apostrophe_is_safe_for_history_links(self):
        filename = chat_history.save_chat("Sam's chat", [{"user": "hi", "assistant": "hello"}])

        self.assertNotIn("'", filename)
        self.assertTrue(filename.endswith(".json"))
        self.assertEqual(chat_history.load_chat(filename)[0]["user"], "hi")

    def test_legacy_json_files_are_removed_after_migration(self):
        legacy_path = os.path.join(self.temp_dir, "legacy_chat.json")
        with open(legacy_path, "w", encoding="utf-8") as file:
            json.dump([{"user": "old", "assistant": "value"}], file)

        chat_history.init_db()

        self.assertEqual(chat_history.list_history(), ["legacy_chat.json"])
        self.assertFalse(os.path.exists(legacy_path))

    def test_same_conversation_reuses_one_history_row(self):
        conversation_id = "conversation-123"
        messages_1 = [{"user": "hi", "assistant": "hello"}]
        messages_2 = [{"user": "hi", "assistant": "hello"}, {"user": "again", "assistant": "still here"}]

        first_filename = chat_history.save_chat("hello", messages_1, conversation_id)
        second_filename = chat_history.save_chat("hello", messages_2, conversation_id)

        self.assertEqual(first_filename, second_filename)
        self.assertEqual(len(chat_history.list_history()), 1)
        self.assertEqual(len(chat_history.load_chat(first_filename)), 2)

    def test_general_questions_allow_general_knowledge(self):
        self.assertTrue(is_general_question("What is the capital of France?"))
        self.assertFalse(is_general_question("how many are waiting for approval"))
        self.assertFalse(is_general_question("what is the description of ITEM-2"))
        self.assertFalse(is_general_question("which commodity group is ITEM-2 in"))
        self.assertFalse(is_general_question("what medical leave is in the contract"))
        self.assertFalse(is_general_question("what salary and allowance terms apply"))

        with patch("backend.routers.build_file_context", return_value=""):
            prompt = build_prompt("What is the capital of France?")

        lowered = prompt.lower()
        self.assertIn("general knowledge", lowered)
        self.assertIn("ignore this section", lowered)
        self.assertNotIn("no information in the uploaded files", lowered)

    def test_pm_procedure_download_uses_maximo_document_link(self):
        pm_result = {
            "member": [{
                "pmnum": "PM-401",
                "description": "Pump maintenance procedure",
                "doclinks": [{
                    "href": "https://example.maximo.com/maximo/docs/pump-maintenance-procedure.pdf",
                    "title": "Pump Maintenance Procedure"
                }]
            }]
        }

        with patch("backend.routers.get_object", return_value=pm_result):
            answer = find_maximo_document_download(
                "Download the PM procedure for pump maintenance"
            )

        self.assertTrue(answer.startswith("/download?url="))
        self.assertIn("pump-maintenance-procedure.pdf", answer)

    def test_download_requests_select_registered_maximo_exports(self):
        self.assertEqual(
            maximo_download_tool_requested("Download the item list"),
            "download_item_list"
        )
        self.assertEqual(
            maximo_download_tool_requested("Export all work orders as CSV"),
            "download_work_orders"
        )
        self.assertEqual(
            maximo_download_tool_requested("Download assets"),
            "download_assets"
        )
        self.assertEqual(
            maximo_download_tool_requested("Download work order 1001"),
            "download_work_orders"
        )
        self.assertIsNone(maximo_download_tool_requested("List all assets"))

    def test_summary_intent_fetches_maximo_before_calling_qwen(self):
        events = []
        question = "summarize work order 1001"
        target = maximo_summary_target(question)
        self.assertEqual(classify_chat_intent(question), "summary")

        def fetch_work_order(wonum):
            events.append(("maximo", wonum))
            return {"member": [{"wonum": wonum, "status": "WAPPR"}]}

        def summarize(messages):
            events.append(("qwen", messages[0]["content"]))
            return "Work order 1001 is waiting for approval."

        with patch("backend.routers.get_work_order", side_effect=fetch_work_order), \
             patch("backend.routers.generate_response", side_effect=summarize):
            answer = get_maximo_summary_answer(question, target)

        self.assertEqual(answer, "Work order 1001 is waiting for approval.")
        self.assertEqual(events[0], ("maximo", "1001"))
        self.assertEqual(events[1][0], "qwen")
        self.assertIn('"status": "WAPPR"', events[1][1])

    def test_summary_selects_only_requested_maximo_fields_and_keeps_identifier(self):
        question = "show asset 11340 location description status"
        target = maximo_summary_target(question)

        self.assertEqual(
            maximo_select_fields(question, target),
            "assetnum,location,description,status"
        )
        self.assertIsNone(
            maximo_select_fields("show asset 11340", target)
        )

    def test_download_it_resolves_the_recent_maximo_record(self):
        history = [
            {"role": "user", "content": "What is the status of work order 1001?"},
            {"role": "assistant", "content": "Work order 1001 is waiting for approval."}
        ]

        self.assertEqual(
            resolve_maximo_download_request("download it", history),
            ("MXAPIWO", "wonum", "1001")
        )
        self.assertEqual(
            resolve_maximo_download_request("download it", []),
            (None, None, None)
        )

    def test_specific_asset_and_item_downloads_select_record_exports(self):
        self.assertEqual(
            maximo_download_tool_requested("Download asset 11340"),
            "download_assets"
        )
        self.assertEqual(
            maximo_download_tool_requested("Download item ITEM-123"),
            "download_item_list"
        )

    def test_generic_resolver_accepts_any_catalogued_maximo_object(self):
        self.assertEqual(
            maximo_download_object_requested("Download MXAPIINCIDENT"),
            "MXAPIINCIDENT"
        )
        self.assertEqual(
            maximo_download_object_requested("Download company master"),
            "MXAPIVENDORMSTR"
        )
        self.assertEqual(
            resolve_maximo_download_request("download incident list"),
            ("MXAPIINCIDENT", None, None)
        )
        self.assertEqual(
            resolve_maximo_download_request("What is work order 1001?"),
            (None, None, None)
        )
        self.assertEqual(
            resolve_maximo_download_request("download tool inventory"),
            ("MXAPITOOLINV", None, None)
        )

    def test_tool_inventory_download_fetches_all_pages_and_returns_excel(self):
        first_page = {
            "member": [{"itemnum": "TOOL-1", "description": "Wrench"}],
            "responseInfo": {
                "nextPage": {"href": "https://maximo.example/api?pageno=2"}
            }
        }
        second_page = {"member": [{"itemnum": "TOOL-2", "latefield": "kept"}]}

        async def read_response(response):
            return b"".join([chunk async for chunk in response.body_iterator])

        with patch(
            "backend.routers.get_object",
            side_effect=[first_page, second_page]
        ) as get_object:
            response = download_maximo_object("MXAPITOOLINV")
            payload = asyncio.run(read_response(response))

        workbook = load_workbook(BytesIO(payload), read_only=True)
        rows = list(workbook["Tool Inventory"].values)
        self.assertIn(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            response.media_type
        )
        self.assertIn('filename="toolinv.xlsx"', response.headers["content-disposition"])
        self.assertEqual(rows[0], ("itemnum", "description", "latefield"))
        self.assertEqual(rows[2], ("TOOL-2", None, "kept"))
        self.assertEqual(get_object.call_count, 2)

    def test_generic_download_exports_scalar_columns_for_unlisted_object(self):
        async def read_response(response):
            return b"".join([chunk async for chunk in response.body_iterator])

        with patch(
            "backend.routers.get_object",
            return_value={"member": [{"ticketid": "INC-1", "description": "Pump issue", "status": "NEW"}]}
        ) as get_object:
            response = download_maximo_object("MXAPIINCIDENT")
            payload = asyncio.run(read_response(response))

        workbook = load_workbook(BytesIO(payload), read_only=True)
        rows = list(workbook["Maximo Data"].values)
        self.assertIn('filename="incident.xlsx"', response.headers["content-disposition"])
        self.assertEqual(rows, [
            ("ticketid", "description", "status"),
            ("INC-1", "Pump issue", "NEW")
        ])
        get_object.assert_called_once_with(
            "MXAPIINCIDENT",
            page_size=50,
            where=None,
            page_number=1
        )

    def test_generic_download_includes_columns_found_on_later_pages(self):
        first_page = {
            "member": [{"wonum": "WO-1", "description": "First"}],
            "responseInfo": {
                "nextPage": {"href": "https://maximo.example/api?pageno=2"}
            }
        }
        second_page = {"member": [{"wonum": "WO-2", "late_field": "preserved"}]}

        async def read_response(response):
            return b"".join([chunk async for chunk in response.body_iterator])

        with patch(
            "backend.routers.get_object",
            side_effect=[first_page, second_page]
        ):
            response = download_maximo_object("MXAPIWO")
            payload = asyncio.run(read_response(response))

        workbook = load_workbook(BytesIO(payload), read_only=True)
        rows = list(workbook["Maximo Data"].values)
        self.assertEqual(rows[0], ("wonum", "description", "late_field"))
        self.assertEqual(rows[2], ("WO-2", None, "preserved"))

    def test_chat_stream_asks_for_target_instead_of_listing_on_bare_download(self):
        async def read_response(response):
            return "".join([chunk async for chunk in response.body_iterator])

        response = chat_stream(ChatRequest(message="download it"))

        self.assertEqual(
            asyncio.run(read_response(response)),
            "Which Maximo object or record should I download?"
        )

    def test_download_tool_returns_paginated_maximo_csv_attachment(self):
        first_page = {
            "member": [{
                "assetnum": "A-1",
                "description": "Pump",
                "priority": 0,
                "status": "ACTIVE"
            }],
            "responseInfo": {
                "nextPage": {"href": "https://maximo.example/api?pageno=2"}
            }
        }
        second_page = {"member": [{"assetnum": "A-2", "status": "OPERATING"}]}

        async def read_response(response):
            return "".join([chunk async for chunk in response.body_iterator])

        with patch("backend.routers.get_object", side_effect=[first_page, second_page]) as get_object:
            response = download("download_assets")
            body = asyncio.run(read_response(response))

        self.assertEqual(response.media_type, "text/csv")
        self.assertIn('filename="assets.csv"', response.headers["content-disposition"])
        self.assertTrue(body.startswith("assetnum,description,priority,status"))
        self.assertIn("A-1,Pump,0,ACTIVE", body)
        self.assertIn("A-2,,,OPERATING", body)
        self.assertNotIn("_additional_fields", body)
        self.assertNotIn("href", body)
        rows = list(csv.DictReader(StringIO(body)))
        self.assertEqual(rows[0]["priority"], "0")
        self.assertEqual(get_object.call_count, 2)
        self.assertEqual(get_object.call_args_list[1].kwargs["page_number"], 2)

    def test_chat_stream_returns_registered_download_route(self):
        async def read_response(response):
            return "".join([chunk async for chunk in response.body_iterator])

        response = chat_stream(ChatRequest(message="Download assets"))

        self.assertEqual(
            asyncio.run(read_response(response)),
            "/download/maximo/MXAPIASSET"
        )

    def test_chat_stream_routes_tool_inventory_to_excel_endpoint(self):
        async def read_response(response):
            return "".join([chunk async for chunk in response.body_iterator])

        response = chat_stream(ChatRequest(message="download tool inventory"))

        self.assertEqual(
            asyncio.run(read_response(response)),
            "/download/maximo/MXAPITOOLINV"
        )

    def test_chat_stream_passes_specific_work_order_to_download_tool(self):
        async def read_response(response):
            return "".join([chunk async for chunk in response.body_iterator])

        response = chat_stream(ChatRequest(message="Download work order 1001"))

        self.assertEqual(
            asyncio.run(read_response(response)),
            "/download/maximo/MXAPIWO?where_field=wonum&identifier=1001"
        )

    def test_chat_stream_resolves_download_it_from_recent_history(self):
        async def read_response(response):
            return "".join([chunk async for chunk in response.body_iterator])

        response = chat_stream(ChatRequest(
            message="download it",
            history=[
                {"role": "user", "content": "What is the status of work order 1001?"},
                {"role": "assistant", "content": "Work order 1001 is waiting for approval."}
            ]
        ))

        self.assertEqual(
            asyncio.run(read_response(response)),
            "/download/maximo/MXAPIWO?where_field=wonum&identifier=1001"
        )

    def test_work_order_download_filters_maximo_and_names_csv(self):
        async def read_response(response):
            return "".join([chunk async for chunk in response.body_iterator])

        with patch(
            "backend.routers.get_object",
            return_value={"member": [{"wonum": "1001", "status": "WAPPR"}]}
        ) as get_object:
            response = download("download_work_orders", identifier="1001")
            body = asyncio.run(read_response(response))

        self.assertIn('filename="work_order_1001.csv"', response.headers["content-disposition"])
        self.assertIn("wonum,status", body)
        self.assertIn("1001,WAPPR", body)
        self.assertEqual(get_object.call_args.kwargs["where"], 'wonum="1001"')

    def test_non_pm_download_questions_do_not_trigger_file_downloads(self):
        self.assertFalse(
            maximo_document_download_requested("Open IBM and show the company page")
        )
        self.assertFalse(
            maximo_document_download_requested("Download the latest pricing sheet")
        )

    def test_non_document_homepage_urls_are_not_used_for_downloads(self):
        self.assertFalse(
            maximo_document_download_requested("Download the MRO homepage")
        )
        self.assertFalse(_is_downloadable_document_url("http://www.mro.com"))
        self.assertFalse(_is_downloadable_document_url("https://www.mro.com"))

    def test_work_order_prompt_uses_maximo_without_files_or_history(self):
        with patch("backend.routers.core.CHAT_SUMMARY", "old conversation summary"), \
             patch("backend.routers.RECENT_MESSAGES", [{"user": "old conversation"}]), \
             patch("backend.routers.build_file_context") as file_context, \
             patch("backend.routers.get_work_order", return_value={"member": [{"wonum": "1001"}]}) as get_work_order:
            prompt = build_prompt("What is the status of WO 1001?")

        self.assertEqual(maximo_work_order_number("WO 1001"), "1001")
        get_work_order.assert_called_once_with("1001")
        file_context.assert_not_called()
        self.assertIn('"wonum": "1001"', prompt)
        self.assertNotIn("old conversation", prompt)

    def test_asset_prompt_uses_maximo_without_files_or_history(self):
        with patch("backend.routers.core.CHAT_SUMMARY", "old conversation summary"), \
             patch("backend.routers.RECENT_MESSAGES", [{"user": "old conversation"}]), \
             patch("backend.routers.build_file_context") as file_context, \
             patch("backend.routers.get_asset", return_value={"member": [{"assetnum": "11340"}]}) as get_asset:
            prompt = build_prompt("What is the status of asset 11340?")

        self.assertEqual(maximo_asset_number("asset 11340"), "11340")
        get_asset.assert_called_once_with("11340")
        file_context.assert_not_called()
        self.assertIn('"assetnum": "11340"', prompt)
        self.assertNotIn("old conversation", prompt)

    def test_asset_chat_bypasses_spreadsheet_and_clears_memory(self):
        recent_messages = [{"user": "old conversation"}]
        asset_result = {"member": [{
            "assetnum": "11340",
            "description": "Motor starter",
            "status_description": "Not Ready",
            "siteid": "BEDFORD",
            "location": "BR300"
        }]}
        with patch("backend.routers.RECENT_MESSAGES", recent_messages), \
             patch("backend.routers.core.CHAT_SUMMARY", "old summary"), \
             patch("backend.routers.build_spreadsheet_answer") as spreadsheet_answer, \
             patch("backend.routers.get_asset", return_value=asset_result) as get_asset, \
             patch("backend.routers.build_prompt") as build_prompt, \
             patch("backend.routers.generate_response", return_value="Asset 11340 is not ready at BEDFORD.") as generate_response:
            response = chat(ChatRequest(message="Asset 11340"))
            self.assertEqual(recent_messages, [])

        spreadsheet_answer.assert_not_called()
        get_asset.assert_called_once_with("11340")
        build_prompt.assert_not_called()
        generate_response.assert_called_once()
        self.assertEqual(
            response["response"],
            "Asset 11340 is not ready at BEDFORD."
        )

    def test_asset_streaming_chat_returns_maximo_summary_directly(self):
        asset_result = {"member": [{
            "assetnum": "11340",
            "status_description": "Not Ready"
        }]}
        stream_capture = {}

        def capture_stream(content, media_type):
            stream_capture["content"] = list(content)
            stream_capture["media_type"] = media_type
            return stream_capture

        with patch("backend.routers.build_spreadsheet_answer") as spreadsheet_answer, \
             patch("backend.routers.get_asset", return_value=asset_result) as get_asset, \
             patch("backend.routers.build_prompt") as build_prompt, \
             patch("backend.routers.generate_response", return_value="Asset 11340 is not ready."), \
             patch("backend.routers.llm.create_chat_completion_stream") as model_stream, \
             patch("backend.routers.StreamingResponse", side_effect=capture_stream):
            response = chat_stream(ChatRequest(message="Asset 11340"))

        spreadsheet_answer.assert_not_called()
        get_asset.assert_called_once_with("11340")
        build_prompt.assert_not_called()
        model_stream.assert_not_called()
        self.assertEqual(
            response["content"],
            ["Asset 11340 is not ready."]
        )
        self.assertEqual(response["media_type"], "text/plain")

    def test_contract_and_tool_queries_are_detected_as_maximo_lookups(self):
        self.assertEqual(maximo_reference("show contract CT-123"), ("contract", "CT-123"))
        self.assertEqual(maximo_reference("find tool TOOL-45"), ("tool", "TOOL-45"))
        self.assertEqual(maximo_reference("find tool VACUUM"), ("tool", "VACUUM"))
        self.assertEqual(maximo_reference("show purchase order 45001"), ("purchase order", "45001"))
        self.assertEqual(maximo_reference("what tools are available"), (None, None))

    def test_explicit_maximo_object_query_is_parsed(self):
        self.assertEqual(
            maximo_object_query("MXAPIITEM itemnum=ITEM-123"),
            ("MXAPIITEM", "itemnum", "ITEM-123")
        )

    def test_collection_queries_route_plural_resources_to_maximo(self):
        self.assertEqual(maximo_collection_query("show tool records"), "MXAPITOOLINV")
        self.assertEqual(maximo_collection_query("what tools are available"), "MXAPITOOLINV")
        self.assertEqual(maximo_collection_query("how many tools do we have"), "MXAPITOOLINV")
        self.assertEqual(maximo_collection_query("show purchase order records"), "MXAPIPO")
        self.assertEqual(maximo_collection_query("list purchase orders"), "MXAPIPO")
        self.assertIsNone(maximo_collection_query("purchase order 20004"))

    def test_collection_count_intent_routes_broad_count_questions(self):
        self.assertEqual(maximo_collection_query("how many work orders exist"), "MXAPIWO")
        self.assertEqual(maximo_collection_query("total assets"), "MXAPIASSET")
        self.assertTrue(maximo_collection_count_requested("how many tools are active"))
        self.assertTrue(maximo_collection_count_requested("total number of assets"))
        self.assertFalse(maximo_collection_count_requested("show tools"))

    def test_streaming_count_routes_to_maximo_before_spreadsheet_answer(self):
        capture = {}

        def capture_stream(content, media_type):
            capture["content"] = list(content)
            capture["media_type"] = media_type
            return capture

        with patch("backend.routers.build_spreadsheet_answer") as spreadsheet_answer, \
             patch("backend.routers.get_maximo_chat_answer", return_value="Maximo returned 123 tools matching your query.") as get_maximo_answer, \
             patch("backend.routers.StreamingResponse", side_effect=capture_stream):
            response = chat_stream(ChatRequest(message="how many tools are there"))

        spreadsheet_answer.assert_not_called()
        get_maximo_answer.assert_called_once_with(
            "how many tools are there",
            collection_object="MXAPITOOLINV"
        )
        self.assertEqual(
            response["content"],
            ["Maximo returned 123 tools matching your query."]
        )

    def test_tool_questions_list_records_unless_summary_is_requested(self):
        self.assertFalse(tool_inventory_summary_requested("show tools"))
        self.assertFalse(tool_inventory_summary_requested("what tools are available"))
        self.assertTrue(tool_inventory_summary_requested("summarize the tool inventory"))
        self.assertTrue(tool_inventory_summary_requested("how many tools are active"))

        result = build_tool_inventory_list([
            {
                "itemnum": "TOOL-45",
                "description": "Torque wrench",
                "status": "ACTIVE",
                "siteid": "SITE-X",
                "location": "TOOLROOM"
            }
        ])

        self.assertIn("TOOL-45: Torque wrench", result)
        self.assertIn("Site SITE-X", result)
        self.assertIn("Location TOOLROOM", result)
        self.assertNotIn("Largest concentration", result)

    def test_tool_inventory_summary_reports_status_site_and_location(self):
        records = [
            {"status": "ACTIVE", "siteid": "SITE-X", "location": "TOOLROOM"},
            {"status_description": "Active", "siteid": "SITE-X", "location": "TOOLROOM"},
            {"status": "INACTIVE", "siteid": "SITE-Y", "location": "STORES"},
        ]

        summary = build_tool_inventory_summary(records)

        self.assertIn("3 tool records", summary)
        self.assertIn("Most tools are active: 2 of 3 (67%)", summary)
        self.assertIn("Largest concentration by site: SITE-X (2 tools)", summary)
        self.assertIn("Largest concentration by location: TOOLROOM (2 tools)", summary)

    def test_tool_summary_uses_only_live_tool_api_results(self):
        api_result = {"member": [
            {"status": "ACTIVE", "siteid": "SITE-X", "location": "CENTRAL"},
            {"status": "ACTIVE", "siteid": "SITE-X", "location": "CENTRAL"},
            {"status": "INACTIVE", "siteid": "SITE-Y", "location": "STORES"},
        ]}
        with patch("backend.routers.get_object", return_value=api_result) as get_object, \
             patch("backend.routers.search_saved_history") as search_history, \
             patch("backend.routers.build_file_context") as file_context, \
             patch("backend.routers.generate_response") as generate_response:
            answer = get_maximo_chat_answer(
                "give me a summary on the tools",
                collection_object="MXAPITOOLINV"
            )

        get_object.assert_called_once_with(
            "MXAPITOOLINV",
            page_size=50,
            where=None
        )
        search_history.assert_not_called()
        file_context.assert_not_called()
        generate_response.assert_not_called()
        self.assertIn("3 tool records", answer)
        self.assertIn("Largest concentration by site: SITE-X (2 tools)", answer)
        self.assertNotIn("work orders", answer.lower())

    def test_collection_priority_queries_add_wopriority_filter(self):
        self.assertEqual(
            maximo_collection_where("show work orders with priority 7 and 9", "MXAPIWO"),
            "wopriority=7 or wopriority=9"
        )
        self.assertEqual(
            maximo_collection_where("show work orders with priority 7", "MXAPIWO"),
            "wopriority=7"
        )
        self.assertIsNone(maximo_collection_where("show work orders", "MXAPIWO"))

    def test_purchase_order_collection_chat_queries_maximo_after_document_search(self):
        api_result = {"member": [
            {"ponum": "20004", "description": "Motor parts", "status": "APPR"},
            {"ponum": "20005", "description": "Electrical supplies", "status": "WAPPR"}
        ]}
        search_order = []

        def search_documents(question):
            search_order.append("documents")
            return "No uploaded PO details."

        def query_maximo(object_name, page_size):
            search_order.append("maximo")
            return api_result

        with patch("backend.routers.chat_history.list_history", return_value=[]), \
             patch("backend.routers.build_file_context", side_effect=search_documents), \
             patch("backend.routers.get_object", side_effect=query_maximo) as get_object, \
             patch("backend.routers.generate_response", return_value="Purchase orders 20004 and 20005 were found.") as generate_response:
            response = chat(ChatRequest(message="show purchase order records"))

        self.assertEqual(search_order, ["documents", "maximo"])
        get_object.assert_called_once_with("MXAPIPO", page_size=10)
        prompt = generate_response.call_args.args[0][0]["content"]
        self.assertIn("20004: Motor parts (APPR)", prompt)
        self.assertIn("20005: Electrical supplies (WAPPR)", prompt)
        self.assertEqual(response["response"], "Purchase orders 20004 and 20005 were found.")

    def test_explicit_maximo_object_chat_bypasses_files_and_llm(self):
        result = {"member": [{
            "itemnum": "ITEM-123",
            "description": "Replacement filter",
            "status": "ACTIVE"
        }]}
        with patch("backend.routers.build_spreadsheet_answer") as spreadsheet_answer, \
             patch("backend.routers.get_record", return_value=result) as get_record, \
             patch("backend.routers.build_prompt") as build_prompt, \
             patch("backend.routers.generate_response", return_value="The item is an active replacement filter.") as generate_response:
            response = chat(ChatRequest(message="MXAPIITEM itemnum=ITEM-123"))

        spreadsheet_answer.assert_not_called()
        get_record.assert_called_once_with("MXAPIITEM", "itemnum", "ITEM-123")
        build_prompt.assert_not_called()
        generate_response.assert_called_once()
        prompt = generate_response.call_args.args[0][0]["content"]
        self.assertIn("Replacement filter", prompt)
        self.assertEqual(response["response"], "The item is an active replacement filter.")

    def test_contract_chat_bypasses_uploaded_file_and_llm_paths(self):
        contract_result = {"member": [{
            "contractnum": "CT-123",
            "description": "Maintenance contract",
            "status_description": "Approved",
            "vendor": "ACME"
        }]}
        search_order = []

        def search_saved_chats():
            search_order.append("history")
            return ["saved-chat.json"]

        def load_saved_chat(filename):
            return [{
                "role": "user",
                "content": "We reviewed contract CT-123 renewal last June."
            }]

        def search_documents(question):
            search_order.append("documents")
            return "Uploaded contract note: renewal reviewed."

        def query_contract(identifier):
            search_order.append("maximo")
            return contract_result

        with patch("backend.routers.build_spreadsheet_answer") as spreadsheet_answer, \
               patch("backend.routers.chat_history.list_history", side_effect=search_saved_chats), \
               patch("backend.routers.chat_history.load_chat", side_effect=load_saved_chat), \
             patch("backend.routers.build_file_context", side_effect=search_documents), \
             patch("backend.routers.get_contract", side_effect=query_contract) as get_contract, \
             patch("backend.routers.build_prompt") as build_prompt, \
             patch(
                 "backend.routers.generate_response",
                 return_value="Contract CT-123 is approved; renewal was reviewed last June."
             ) as generate_response:
            response = chat(ChatRequest(message="show contract CT-123"))

        spreadsheet_answer.assert_not_called()
        get_contract.assert_called_once_with("CT-123")
        self.assertEqual(search_order, ["history", "documents", "maximo"])
        build_prompt.assert_not_called()
        generate_response.assert_called_once()
        prompt = generate_response.call_args.args[0][0]["content"]
        self.assertIn("We reviewed contract CT-123 renewal last June", prompt)
        self.assertIn("Uploaded contract note: renewal reviewed", prompt)
        self.assertIn("Contract CT-123 is Approved", prompt)
        self.assertEqual(
            response["response"],
            "Contract CT-123 is approved; renewal was reviewed last June."
        )
        self.assertNotIn("From saved chat history", response["response"])
        self.assertNotIn("From live Maximo API", response["response"])

    def test_tool_chat_bypasses_uploaded_file_and_llm_paths(self):
        tool_result = {"member": [{
            "itemnum": "TOOL-45",
            "description": "Insulated wrench",
            "status": "ACTIVE",
            "location": "TOOLROOM"
        }]}
        with patch("backend.routers.build_spreadsheet_answer") as spreadsheet_answer, \
             patch("backend.routers.get_tool", return_value=tool_result) as get_tool, \
             patch("backend.routers.build_prompt") as build_prompt, \
             patch("backend.routers.generate_response", return_value="Tool TOOL-45 is an active insulated wrench in TOOLROOM.") as generate_response:
            response = chat(ChatRequest(message="tool TOOL-45"))

        spreadsheet_answer.assert_not_called()
        get_tool.assert_called_once_with("TOOL-45")
        build_prompt.assert_not_called()
        generate_response.assert_called_once()
        self.assertEqual(
            response["response"],
            "Tool TOOL-45 is an active insulated wrench in TOOLROOM."
        )

    def test_work_order_chat_bypasses_spreadsheet_and_clears_memory(self):
        recent_messages = [{"user": "old conversation"}]
        maximo_result = {"member": [{
            "wonum": "1001",
            "description": "12 Month Service on Shipping Dept #1 Conveyor",
            "status": "INPRG",
            "status_description": "In progress",
            "siteid": "BEDFORD",
            "location": "SHIPPING"
        }]}
        with patch("backend.routers.RECENT_MESSAGES", recent_messages), \
             patch("backend.routers.core.CHAT_SUMMARY", "old summary"), \
             patch("backend.routers.build_spreadsheet_answer") as spreadsheet_answer, \
             patch("backend.routers.get_work_order", return_value=maximo_result) as get_work_order, \
             patch("backend.routers.build_prompt") as build_prompt, \
             patch("backend.routers.generate_response", return_value="Work order 1001 is in progress for the conveyor.") as generate_response:
            response = chat(ChatRequest(message="WO 1001"))
            self.assertEqual(recent_messages, [])

        spreadsheet_answer.assert_not_called()
        get_work_order.assert_called_once_with("1001")
        build_prompt.assert_not_called()
        generate_response.assert_called_once()
        self.assertEqual(
            response["response"],
            "Work order 1001 is in progress for the conveyor."
        )

    def test_streaming_work_order_chat_bypasses_spreadsheet_answers(self):
        maximo_result = {"member": [{
            "wonum": "1001",
            "status_description": "In progress"
        }]}
        stream_capture = {}

        def capture_stream(content, media_type):
            stream_capture["content"] = list(content)
            stream_capture["media_type"] = media_type
            return stream_capture

        with patch("backend.routers.RECENT_MESSAGES", [{"user": "old conversation"}]), \
             patch("backend.routers.core.CHAT_SUMMARY", "old summary"), \
             patch("backend.routers.build_spreadsheet_answer") as spreadsheet_answer, \
             patch("backend.routers.get_work_order", return_value=maximo_result) as get_work_order, \
             patch("backend.routers.build_prompt") as build_prompt, \
             patch("backend.routers.generate_response", return_value="Work order 1001 is in progress."), \
             patch("backend.routers.llm.create_chat_completion_stream") as model_stream, \
             patch("backend.routers.StreamingResponse", side_effect=capture_stream):
            response = chat_stream(ChatRequest(message="WO 1001"))

        spreadsheet_answer.assert_not_called()
        get_work_order.assert_called_once_with("1001")
        build_prompt.assert_not_called()
        model_stream.assert_not_called()
        self.assertEqual(
            response["content"],
            ["Work order 1001 is in progress."]
        )
        self.assertEqual(response["media_type"], "text/plain")

    def test_non_work_order_count_questions_do_not_use_work_order_logic(self):
        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/workbook.csv"]), \
             patch("backend.services.file_service.load_file", return_value="Work Order | Status\nA1 | Open\nA2 | Closed\n"):
            answer = build_spreadsheet_count_answer("how many colours are there")

        self.assertEqual(answer, "")

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/workbook.csv"]), \
               patch("backend.services.file_service.load_file", return_value="Work Order,Status\nA1,Open\nA2,Closed\n"):
            answer = build_spreadsheet_count_answer("how many work orders are open")

        self.assertIn("work orders", answer.lower())

    def test_status_count_questions_use_status_values_from_any_spreadsheet(self):
        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/approvals.csv"]), \
               patch("backend.services.file_service.load_file", return_value="Request,Approval Status\nA1,Approved\nA2,Pending\nA3,Approved\n"):
            answer = build_spreadsheet_count_answer("how many approved")

        self.assertIn("2 approved", answer.lower())
        self.assertIn("approvals.csv", answer)

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/workbook.xlsx"]), \
             patch("backend.services.file_service.load_file", return_value="Work Order | Status\nA1 | APPR\nA2 | WAPPR\nA3 | APPR\n"):
            answer = build_spreadsheet_count_answer("how many approved")

        self.assertIn("2 approved", answer.lower())
        self.assertNotIn("wappr", answer.lower())

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/workbook.xlsx"]), \
             patch("backend.services.file_service.load_file", return_value="Work Order | Status\nA1 | WAPPR\nA2 | APPR\n"):
            answer = build_spreadsheet_status_answer("what does WAPPR mean")

        self.assertIn("waiting approval", answer.lower())

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/workbook.xlsx"]), \
             patch("backend.services.file_service.load_file", return_value="Work Order | Status\nA1 | WAPPR\nA2 | APPR\n"):
            answer = build_spreadsheet_count_answer("how many are waiting for approval")

        self.assertIn("1 waiting approval", answer.lower())

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/WOTRACK.csv"]), \
             patch("backend.services.file_service.load_file", return_value="A1,Task one,,,,,,,,,WAPPR\nA2,Task two,,,,,,,,,APPR\nA3,Task three,,,,,,,,,APPR\n"):
            waiting_answer = build_spreadsheet_count_answer("how many are waiting for approval")
            approved_answer = build_spreadsheet_count_answer("how many are approved")

        self.assertIn("1 waiting approval", waiting_answer.lower())
        self.assertIn("2 approved", approved_answer.lower())

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/WOTRACK1.csv", "/tmp/WOTRACK2.csv"]), \
             patch("backend.services.file_service.load_file", side_effect=[
                 "Work Order,Status\nA1,WAPPR\nA2,WAPPR\nA3,APPR\nA4,APPR\nA5,WAPPR\nA6,WAPPR\nA7,WAPPR\nA8,WAPPR\nA9,WAPPR\nA10,WAPPR\n",
                 "Work Order,Status\nB1,WAPPR\nB2,WAPPR\nB3,WAPPR\nB4,WAPPR\nB5,WAPPR\nB6,WAPPR\nB7,WAPPR\nB8,WAPPR\nB9,WAPPR\nB10,WAPPR\n"
             ]):
            waiting_answer = build_spreadsheet_count_answer("how many are waiting for approval")

        self.assertIn("18 waiting approval", waiting_answer.lower())

    def test_priority_column_questions_use_actual_file_values(self):
        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/priority.csv"]), \
             patch("backend.services.file_service.load_file", return_value="Work Order,Priority,Status\nA1,7,APPR\nA2,9,WAPPR\nA3,9,APPR\n"):
            answer = build_spreadsheet_priority_answer("what is the highest priority")

        self.assertIn("9", answer)
        self.assertIn("highest priority", answer.lower())

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/priority.csv"]), \
             patch("backend.services.file_service.load_file", return_value="Work Order,Priority,Status\nA1,7,APPR\nA2,9,WAPPR\nA3,9,APPR\n"):
            answer = build_spreadsheet_priority_answer("how many work orders have priority 7")

        self.assertIn("1", answer)
        self.assertIn("priority 7", answer.lower())

    def test_file_context_labels_all_spreadsheet_attributes(self):
        filepath = os.path.join(self.temp_dir, "work_orders.csv")
        with open(filepath, "w", encoding="utf-8") as file:
            file.write(
                "Priority,Description,Location,Is dynamic,Work Order,Site,"
                "Organization,Column 8,Scheduled Start,Asset,Status\n"
                "2,Task one,SHIPPING,NO,1001,BEDFORD,497,,,ASSET-1,WAPPR\n"
            )

        documents, metadata = _normalized_tabular_records(filepath)

        self.assertIn("Priority=2", documents[0])
        self.assertIn("Location=SHIPPING", documents[0])
        self.assertIn("Work Order=1001", documents[0])
        self.assertEqual(metadata[0]["attr_site"], "BEDFORD")
        self.assertEqual(metadata[0]["attr_status"], "WAPPR")

    def test_file_context_routes_queries_through_metadata_and_vector_search(self):
        expected_context = "Row 3: Item=ITEM-2 | Description=Filter"
        with patch("backend.services.file_service.search_context", return_value=expected_context) as search:
            context = build_file_context("what is the description of ITEM-2")

        self.assertEqual(context, expected_context)
        search.assert_called_once_with("what is the description of ITEM-2")

    def test_exact_attribute_lookup_uses_sqlite_attribute_filter(self):
        filename = "uploads/1790569151201.xlsx"
        row_metadata = {
            "context_id": "context-1",
            "document_id": os.path.basename(filename),
            "filename": os.path.basename(filename),
            "file_type": "xlsx",
            "chunk_index": 1,
            "record_type": "row",
            "row_index": 1,
            "attr_item": "ITEM-2",
            "attr_description": "Filter",
            "attr_status": "ACTIVE"
        }
        with patch.object(rag_service, "restore_persisted_session"), \
             patch.object(rag_service.core, "SESSION_FILES", [filename]), \
             patch.object(rag_service.core, "ACTIVE_CONTEXT_ID", "context-1"):
            rag_service.index_file(
                os.path.basename(filename),
                "Row 3: Item=ITEM-2 | Description=Filter",
                context_id="context-1",
                documents=["Row 3: Item=ITEM-2 | Description=Filter"],
                document_metadata=[{
                    "row_index": 1,
                    "attr_item": "ITEM-2",
                    "attr_description": "Filter"
                }]
            )
            context = rag_service.search_context("what is the description of ITEM-2")

        self.assertIn("Description=Filter", context)
        connection = rag_service._connect()
        matches = connection.execute(
            "SELECT attribute_value FROM knowledge_attributes WHERE attribute_name = 'item'"
        ).fetchall()
        connection.close()
        self.assertEqual([row[0] for row in matches], ["ITEM-2"])

    def test_semantic_spreadsheet_question_falls_back_to_vector_search(self):
        filename = "uploads/1790569151201.xlsx"
        row_metadata = {
            "context_id": "context-1",
            "document_id": os.path.basename(filename),
            "filename": os.path.basename(filename),
            "file_type": "xlsx",
            "chunk_index": 1,
            "record_type": "row",
            "row_index": 1,
            "attr_item": "ITEM-2",
            "attr_description": "Filter",
            "attr_status": "ACTIVE"
        }
        with patch.object(rag_service, "restore_persisted_session"), \
             patch.object(rag_service.core, "SESSION_FILES", [filename]), \
             patch.object(rag_service.core, "ACTIVE_CONTEXT_ID", "context-1"):
            rag_service.index_file(
                os.path.basename(filename),
                "Row 3: Item=ITEM-2 | Description=Filter for industrial environment",
                context_id="context-1",
                documents=["Row 3: Item=ITEM-2 | Description=Filter for industrial environment"],
                document_metadata=[{
                    "row_index": 1,
                    "attr_item": "ITEM-2",
                    "attr_description": "Filter for industrial environment"
                }]
            )
            context = rag_service.search_context("which item suits an industrial environment")

        self.assertIn("Description=Filter", context)
        self.assertIn("industrial environment", context)

    def test_index_file_persists_column_values_as_metadata(self):
        rag_service.index_file(
            "items.xlsx",
            "Item=ITEM-2 | Description=Filter",
            context_id="context-1",
            documents=["Row 3: Item=ITEM-2 | Description=Filter"],
            document_metadata=[{
                "row_index": 1,
                "attr_item": "ITEM-2",
                "attr_description": "Filter"
            }]
        )

        connection = rag_service._connect()
        stored = connection.execute(
            "SELECT record_type, metadata_json FROM knowledge_documents"
        ).fetchone()
        connection.close()
        stored_metadata = json.loads(stored["metadata_json"])
        self.assertEqual(stored["record_type"], "row")
        self.assertEqual(stored_metadata["attr_item"], "ITEM-2")
        self.assertEqual(stored_metadata["attr_description"], "Filter")

    def test_dataset_summary_uses_all_rows_and_column_values(self):
        metadata_rows = [
            {
                "filename": "inventory.xlsx",
                "record_type": "row",
                "dataset_type": "inventory",
                "column_names": "Item|Commodity Group",
                "attr_item": "ITEM-1",
                "attr_commodity_group": "MECH"
            },
            {
                "filename": "inventory.xlsx",
                "record_type": "row",
                "dataset_type": "inventory",
                "column_names": "Item|Commodity Group",
                "attr_item": "ITEM-2",
                "attr_commodity_group": "TIRE"
            },
            {
                "filename": "inventory.xlsx",
                "record_type": "row",
                "dataset_type": "inventory",
                "column_names": "Item|Commodity Group",
                "attr_item": "ITEM-3",
                "attr_commodity_group": "MECH"
            }
        ]

        summary = rag_service._format_dataset_summaries(metadata_rows)

        self.assertIn("Rows indexed: 3", summary)
        self.assertIn("Commodity Group: 2 distinct values", summary)
        self.assertIn("MECH (2)", summary)
        self.assertIn("TIRE (1)", summary)

    def test_full_dataset_request_returns_every_indexed_row(self):
        filename = "uploads/1790569151201.xlsx"
        row_metadatas = [
            {
                "context_id": "context-1",
                "document_id": os.path.basename(filename),
                "filename": os.path.basename(filename),
                "file_type": "xlsx",
                "chunk_index": row_index,
                "record_type": "row",
                "row_index": row_index,
                "dataset_type": "inventory",
                "column_names": "Item|Commodity Group",
                "attr_item": f"ITEM-{row_index + 1}",
                "attr_commodity_group": "MECH"
            }
            for row_index in range(2)
        ]
        with patch.object(rag_service, "restore_persisted_session"), \
             patch.object(rag_service.core, "SESSION_FILES", [filename]), \
             patch.object(rag_service.core, "ACTIVE_CONTEXT_ID", "context-1"):
            rag_service.index_file(
                os.path.basename(filename),
                " | ".join([
                    "Row 2: Item=ITEM-1",
                    "Row 3: Item=ITEM-2"
                ]),
                context_id="context-1",
                documents=["Row 2: Item=ITEM-1", "Row 3: Item=ITEM-2"],
                document_metadata=[
                    {"row_index": 0, "dataset_type": "inventory", "column_names": "Item"},
                    {"row_index": 1, "dataset_type": "inventory", "column_names": "Item"}
                ]
            )
            context = rag_service.search_context("show the full dataset and all rows")

        self.assertIn("Item=ITEM-1", context)
        self.assertIn("Item=ITEM-2", context)

    def test_small_pdf_search_returns_all_document_chunks(self):
        filename = "uploads/Visagan_Employment Contract (6 mth) (1).pdf"
        chunk_documents = ["Medical Leave: 14 days sick leave.", "Salary: MYR 3,800 monthly."]
        chunk_metadatas = [
            {
                "context_id": "context-1",
                "document_id": os.path.basename(filename),
                "filename": os.path.basename(filename),
                "file_type": "pdf",
                "chunk_index": index,
                "record_type": "chunk",
                "row_index": -1
            }
            for index in range(len(chunk_documents))
        ]
        with patch.object(rag_service, "restore_persisted_session"), \
             patch.object(rag_service.core, "SESSION_FILES", [filename]), \
             patch.object(rag_service.core, "ACTIVE_CONTEXT_ID", "context-1"):
            rag_service.index_file(
                os.path.basename(filename),
                "\n\n".join(chunk_documents),
                context_id="context-1",
                documents=chunk_documents
            )
            context = rag_service.search_context("what salary and leave are included")

        self.assertIn("14 days sick leave", context)
        self.assertIn("MYR 3,800 monthly", context)

    def test_asset_context_labels_asset_attributes(self):
        filepath = os.path.join(self.temp_dir, "assets.csv")
        with open(filepath, "w", encoding="utf-8") as file:
            file.write(
                "Asset,Description,Location,Loop Location,Parent,Rotating Item,"
                "Is M&TE,Linear,Calibration,Site\n"
                "ASSET-1,Pump,SHIPPING,LOOP-1,PARENT-1,YES,NO,YES,2025-01-01,BEDFORD\n"
            )

        documents, metadata = _normalized_tabular_records(filepath)

        self.assertIn("Asset=ASSET-1", documents[0])
        self.assertEqual(metadata[0]["attr_loop_location"], "LOOP-1")
        self.assertEqual(metadata[0]["attr_parent"], "PARENT-1")
        self.assertEqual(metadata[0]["attr_rotating_item"], "Yes")
        self.assertEqual(metadata[0]["attr_is_m_te"], "No")
        self.assertEqual(metadata[0]["attr_calibration"], "2025-01-01")

    def test_asset_count_questions_count_distinct_assets(self):
        csv_text = (
            "2,Task one,SHIPPING,NO,1001,BEDFORD,497,,,ASSET-1,\n"
            "2,Task two,SHIPPING,NO,1002,BEDFORD,497,,,ASSET-2,\n"
            "2,Task three,SHIPPING,NO,1003,BEDFORD,497,,,ASSET-1,\n"
        )
        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/WOTRACK.csv"]), \
             patch("backend.services.file_service.load_file", return_value=csv_text):
            answer = build_spreadsheet_asset_count_answer("how many assets are there")

        self.assertIn("2 distinct assets", answer)

    def test_location_count_questions_use_the_location_column(self):
        csv_text = (
            "2,Task one,HWSTOCK,NO,1001,BEDFORD,497,,,ASSET-1,\n"
            "2,Task two,HWSTOCK,NO,1002,BEDFORD,497,,,ASSET-2,\n"
            "2,Task three,SHIPPING,NO,1003,BEDFORD,497,,,ASSET-3,\n"
        )
        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/WOTRACK.csv"]), \
             patch("backend.services.file_service.load_file", return_value=csv_text):
            answer = build_spreadsheet_location_count_answer("how many location is HWSTOCK")

        self.assertIn("HWSTOCK appears 2 times", answer)

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/WOTRACK.csv"]), \
             patch("backend.services.file_service.load_file", return_value=csv_text):
            answer = build_spreadsheet_distinct_location_count_answer(
                "how many different locations are there"
            )

        self.assertIn("2 different locations", answer)

    def test_asset_filter_count_uses_description(self):
        csv_text = (
            "Asset,Description,Location\n"
            "HD-1,Hard Drive,HWSTOCK\n"
            "HD-2,Hard Drive,HWSTOCK\n"
            "PC-1,Laptop,HARDWARE\n"
        )
        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/assets.csv"]), \
             patch("backend.services.file_service.load_file", return_value=csv_text):
            answer = build_spreadsheet_asset_filter_count_answer("how many of them are hardrive")

        self.assertIn("2 hard-drive assets", answer)

        unix_csv = (
            "Asset,Description,Location\n"
            "A8004,Unix Server,COMP310\n"
            "A8005,Unix Server,COMPLAB\n"
            "A8006,Unix Server,COMP310\n"
        )
        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/assets.csv"]), \
             patch("backend.services.file_service.load_file", return_value=unix_csv):
            answer = build_spreadsheet_asset_filter_count_answer("how many unix server assets are there")

        self.assertIn("3 Unix Server assets", answer)

    def test_inventory_item_count_uses_item_column(self):
        csv_text = (
            "Item,Description,Status\n"
            "ITEM-1,Tire,ACTIVE\n"
            "ITEM-2,Filter,ACTIVE\n"
            "ITEM-1,Tire,ACTIVE\n"
        )
        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/items.csv"]), \
             patch("backend.services.file_service.load_file", return_value=csv_text):
            answer = build_spreadsheet_item_count_answer("how many inventory items are there")

        self.assertIn("2 distinct inventory items", answer)

    def test_distinct_column_count_handles_plural_column_name(self):
        filepath = os.path.join(self.temp_dir, "inventory.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Item", "Commodity Group"])
        sheet.append(["ITEM-1", "TIRE"])
        sheet.append(["ITEM-2", "HARDWARE"])
        sheet.append(["ITEM-3", "TIRE"])
        workbook.save(filepath)

        with patch("backend.services.file_service.core.SESSION_FILES", [filepath]):
            answer = build_spreadsheet_distinct_column_count_answer(
                "how many different commodity groups are there"
            )

        self.assertIn("2 different Commodity Group values", answer)

    def test_relationship_count_question_uses_excel_header(self):
        filepath = os.path.join(self.temp_dir, "relationships.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Relationship", "Description"])
        sheet.append(["REL-1", "First relationship"])
        sheet.append(["REL-2", "Second relationship"])
        workbook.save(filepath)

        with patch("backend.services.file_service.core.SESSION_FILES", [filepath]):
            self.assertFalse(is_general_question("how many relationships are there"))
            answer = build_spreadsheet_column_count_answer(
                "how many relationships are there"
            )

        self.assertIn("2 different Relationship values", answer)

    def test_relationship_type_question_lists_all_matching_names_and_descriptions(self):
        filepath = os.path.join(self.temp_dir, "relationships.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Relationship", "Description", "Type", "Use With"])
        sheet.append(["INCLUDES", "Contains assets", "UNIDIRECTIONAL", "ASSET"])
        sheet.append(["CONNECTS", "Source connects to target", "BIDIRECTIONAL", "ASSET"])
        sheet.append(["RUNSON", "Runs on", "UNIDIRECTIONAL", "CI"])
        sheet.append(["INTERSECTS", "Intersects with", "BIDIRECTIONAL", "CI"])
        workbook.save(filepath)

        with patch("backend.services.file_service.core.SESSION_FILES", [filepath]):
            answer = build_spreadsheet_relationship_list_answer(
                "which relationship is bi and uni, list them"
            )

        self.assertIn("Bidirectional (2): CONNECTS, INTERSECTS", answer)
        self.assertIn("Unidirectional (2): INCLUDES, RUNSON", answer)
        self.assertNotIn("Source connects to target", answer)
        self.assertNotIn("Contains assets", answer)

    def test_relationship_question_filters_and_groups_by_use_with(self):
        filepath = os.path.join(self.temp_dir, "relationships.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Relationship", "Description", "Type", "Use With"])
        sheet.append(["INCLUDES", "Contains assets", "UNIDIRECTIONAL", "ASSET"])
        sheet.append(["CONNECTS", "Connects assets", "BIDIRECTIONAL", "ASSET"])
        sheet.append(["RUNSON", "Runs on", "UNIDIRECTIONAL", "CI"])
        sheet.append(["AFFECTS", "Affects CI", "UNIDIRECTIONAL", "CI"])
        workbook.save(filepath)

        with patch("backend.services.file_service.core.SESSION_FILES", [filepath]):
            ci_answer = build_spreadsheet_relationship_list_answer(
                "which relationships are for CI"
            )
            both_answer = build_spreadsheet_relationship_list_answer(
                "which are CI asset"
            )

        self.assertIn("CI: Unidirectional (2): RUNSON, AFFECTS", ci_answer)
        self.assertNotIn("INCLUDES", ci_answer)
        self.assertNotIn("Runs on", ci_answer)
        self.assertIn("CI: Unidirectional", both_answer)
        self.assertIn("ASSET: Bidirectional", both_answer)

    def test_relationship_list_handler_is_used_by_spreadsheet_answer_router(self):
        filepath = os.path.join(self.temp_dir, "relationships.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Relationship", "Description", "Type"])
        sheet.append(["INCLUDES", "Contains assets", "UNIDIRECTIONAL"])
        sheet.append(["CONNECTS", "Connects assets", "BIDIRECTIONAL"])
        workbook.save(filepath)

        with patch("backend.services.file_service.core.SESSION_FILES", [filepath]):
            answer = build_spreadsheet_answer("which relationship is bi and uni, list them")

        self.assertIn("INCLUDES", answer)
        self.assertIn("CONNECTS", answer)

    def test_generic_column_count_distinguishes_types_and_filters_status(self):
        filepath = os.path.join(self.temp_dir, "contracts.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Contract", "Status", "Type"])
        sheet.append(["C1", "REVISD", "PRICE"])
        sheet.append(["C2", "REVISD", "PRICE"])
        sheet.append(["C3", "REVISD", "PURCHASE"])
        sheet.append(["C4", "APPR", "PURCHASE"])
        workbook.save(filepath)

        with patch("backend.services.file_service.core.SESSION_FILES", [filepath]):
            type_answer = build_spreadsheet_column_count_answer(
                "how many types are there"
            )
            revised_answer = build_spreadsheet_column_count_answer(
                "how many are revised status"
            )

        self.assertIn("2 different Type values", type_answer)
        self.assertIn("3 records with Status revised", revised_answer)

        with patch("backend.services.file_service.core.SESSION_FILES", ["/tmp/WOTRACK.csv"]), \
             patch("backend.services.file_service.load_file", return_value="2,Task one,SHIPPING,NO,1,BEDFORD,497,,,WAPPR\n2,Task two,NEEDHAM,NO,2,BEDFORD,497,,,APPR\n2,Task three,SHIPPING,NO,3,BEDFORD,497,,,APPR\n"):
            answer = build_spreadsheet_location_answer("which location is the most used")

        self.assertIn("most used location is SHIPPING", answer)
        self.assertIn("2 work orders", answer)


if __name__ == "__main__":
    unittest.main()
