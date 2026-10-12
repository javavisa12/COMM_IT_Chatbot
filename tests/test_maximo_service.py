import unittest
from contextlib import redirect_stdout
from io import StringIO
from unittest.mock import Mock, patch

from maximo_service import (
    BASE_URL,
    MAXIMO_OBJECTS,
    MaximoConfigurationError,
    get_object,
    get_asset,
    get_contract,
    get_object_count,
    get_tool,
    get_work_order
)


class MaximoServiceTests(unittest.TestCase):
    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_object_count_follows_pages_to_exact_total(self, mock_get):
        first_page = Mock(status_code=200)
        first_page.json.return_value = {
            "member": [{"itemnum": f"TOOL-{index}"} for index in range(50)],
            "responseInfo": {
                "pagenum": 1,
                "nextPage": {
                    "href": f"{BASE_URL}/api/os/MXAPITOOLINV?pageno=2"
                }
            }
        }
        second_page = Mock(status_code=200)
        second_page.json.return_value = {
            "member": [{"itemnum": "TOOL-50"}, {"itemnum": "TOOL-51"}],
            "responseInfo": {"pagenum": 2}
        }
        mock_get.side_effect = [first_page, second_page]

        self.assertEqual(get_object_count("MXAPITOOLINV"), 52)
        self.assertEqual(mock_get.call_count, 2)
        self.assertEqual(
            mock_get.call_args_list[1].kwargs["params"]["pageno"],
            2
        )

    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_object_accepts_catalogued_structure(self, mock_get):
        mock_response = Mock()
        mock_response.json.return_value = {"member": []}
        mock_get.return_value = mock_response

        get_object("MXAPIITEM", page_size=10)

        self.assertIn("MXAPIITEM", MAXIMO_OBJECTS)
        args, kwargs = mock_get.call_args
        self.assertEqual(args[0], f"{BASE_URL}/api/os/MXAPIITEM")
        self.assertEqual(kwargs["params"]["oslc.pageSize"], 10)

    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_object_sends_requested_oslc_select_fields(self, mock_get):
        mock_response = Mock(status_code=200)
        mock_response.json.return_value = {"member": []}
        mock_get.return_value = mock_response

        get_object(
            "MXAPIASSET",
            page_size=5,
            select="assetnum,location,description,status"
        )

        self.assertEqual(
            mock_get.call_args.kwargs["params"]["oslc.select"],
            "assetnum,location,description,status"
        )

    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_object_rejects_invalid_select_fields(self, mock_get):
        with self.assertRaises(ValueError):
            get_object("MXAPIASSET", select="assetnum, status or 1=1")

        mock_get.assert_not_called()

    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_object_rejects_unknown_structures_and_oversized_pages(self, mock_get):
        with self.assertRaises(ValueError):
            get_object("MXAPI_NOT_IN_CATALOG")
        with self.assertRaises(ValueError):
            get_object("MXAPIITEM", page_size=101)

        mock_get.assert_not_called()

    @patch("maximo_service.API_KEY", "test-secret-key")
    @patch("maximo_service.requests.get")
    def test_request_logging_shows_call_without_api_key(self, mock_get):
        mock_response = Mock(status_code=200)
        mock_response.json.return_value = {"member": [{"itemnum": "ITEM-1"}]}
        mock_get.return_value = mock_response
        output = StringIO()

        with redirect_stdout(output):
            get_object("MXAPIITEM", page_size=1, where='itemnum="ITEM-1"')

        log = output.getvalue()
        self.assertIn("[Maximo API] GET", log)
        self.assertIn('itemnum="ITEM-1"', log)
        self.assertIn("HTTP 200", log)
        self.assertNotIn("test-secret-key", log)

    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_work_order_requests_matching_maximo_record(self, mock_get):
        mock_response = Mock()
        mock_response.json.return_value = {"member": [{"wonum": "1001"}]}
        mock_get.return_value = mock_response

        result = get_work_order("1001")

        self.assertEqual(result, {"member": [{"wonum": "1001"}]})
        mock_get.assert_called_once()
        args, kwargs = mock_get.call_args
        self.assertEqual(args[0], f"{BASE_URL}/api/os/MXAPIWO")
        self.assertEqual(kwargs["params"], {
            "lean": "1",
            "oslc.pageSize": 1,
            "oslc.select": "*",
            "oslc.where": 'wonum="1001"'
        })
        self.assertEqual(kwargs["timeout"], 30)

    @patch("maximo_service.API_KEY", "test-key")
    def test_get_work_order_rejects_invalid_number_without_requesting_api(self):
        with patch("maximo_service.requests.get") as mock_get:
            with self.assertRaises(ValueError):
                get_work_order('1001" or status="CLOSED')

        mock_get.assert_not_called()

    @patch("maximo_service.API_KEY", "")
    def test_get_work_order_requires_api_key(self):
        with patch("maximo_service.requests.get") as mock_get:
            with self.assertRaises(MaximoConfigurationError):
                get_work_order("1001")

        mock_get.assert_not_called()

    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_asset_requests_matching_maximo_asset(self, mock_get):
        mock_response = Mock()
        mock_response.json.return_value = {"member": [{"assetnum": "11340"}]}
        mock_get.return_value = mock_response

        result = get_asset("11340")

        self.assertEqual(result, {"member": [{"assetnum": "11340"}]})
        args, kwargs = mock_get.call_args
        self.assertEqual(args[0], f"{BASE_URL}/api/os/MXAPIASSET")
        self.assertEqual(kwargs["params"], {
            "lean": "1",
            "oslc.pageSize": 1,
            "oslc.select": "*",
            "oslc.where": 'assetnum="11340"'
        })

    @patch("maximo_service.API_KEY", "test-key")
    def test_get_asset_rejects_invalid_number_without_requesting_api(self):
        with patch("maximo_service.requests.get") as mock_get:
            with self.assertRaises(ValueError):
                get_asset('11340" or status="OPERATING')

        mock_get.assert_not_called()

    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_contract_queries_contract_number(self, mock_get):
        mock_response = Mock()
        mock_response.json.return_value = {"member": [{"contractnum": "CT-123"}]}
        mock_get.return_value = mock_response

        result = get_contract("CT-123")

        self.assertEqual(result["member"][0]["contractnum"], "CT-123")
        args, kwargs = mock_get.call_args
        self.assertEqual(args[0], f"{BASE_URL}/api/os/MXAPICONTRACT")
        self.assertEqual(kwargs["params"]["oslc.where"], 'contractnum="CT-123"')

    @patch("maximo_service.API_KEY", "test-key")
    @patch("maximo_service.requests.get")
    def test_get_tool_queries_inventory_item_number(self, mock_get):
        mock_response = Mock()
        mock_response.json.return_value = {"member": [{"itemnum": "TOOL-123"}]}
        mock_get.return_value = mock_response

        result = get_tool("TOOL-123")

        self.assertEqual(result["member"][0]["itemnum"], "TOOL-123")
        args, kwargs = mock_get.call_args
        self.assertEqual(args[0], f"{BASE_URL}/api/os/MXAPITOOLINV")
        self.assertEqual(kwargs["params"]["oslc.where"], 'itemnum="TOOL-123"')