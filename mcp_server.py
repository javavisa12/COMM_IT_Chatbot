from mcp.server.fastmcp import FastMCP
import maximo_service
from maximo_service import (
    get_asset,
    get_work_order,
    get_contract,
    get_tool,
    get_item,
    get_location,
    get_person,
    get_pm,
    get_master_pm,
    get_object,
    get_object_count
)

mcp = FastMCP("Maximo Assistant")


@mcp.tool()
def search_asset(assetnum: str):
    return get_asset(assetnum)

@mcp.tool()
def change_tool_status(
    itemnum: str | None = None,
    status: str = "NOTACTIVE"
):
    return maximo_service.change_tool_status(
        itemnum=itemnum,
        status=status
    )

@mcp.tool()
def search_work_order(wonum: str):
    return get_work_order(wonum)


@mcp.tool()
def search_contract(contractnum: str):
    return get_contract(contractnum)


@mcp.tool()
def search_tool(itemnum: str):
    return get_tool(itemnum)


@mcp.tool()
def search_inventory_item(itemnum: str):
    return get_item(itemnum)


@mcp.tool()
def search_location(location: str):
    return get_location(location)


@mcp.tool()
def search_person(personid: str):
    return get_person(personid)


@mcp.tool()
def search_pm(pmnum: str):
    return get_pm(pmnum)


@mcp.tool()
def search_master_pm(pmnum: str):
    return get_master_pm(pmnum)


@mcp.tool()
def query_maximo_object(
    object_name: str,
    page_size: int = 10,
    where: str | None = None,
    select: str | None = None,
    order_by: str | None = None
):
    return get_object(
        object_name=object_name,
        page_size=page_size,
        where=where,
        select=select,
        order_by=order_by
    )


@mcp.tool()
def count_maximo_records(
    object_name: str,
    where: str | None = None
):
    return {
        "object": object_name,
        "count": get_object_count(
            object_name,
            where=where
        )
    }

@mcp.tool()
def create_work_order(
    description: str,
    siteid: str,
    location: str = None,
    assetnum: str = None
):
    return maximo_service.create_work_order(
        description,
        siteid,
        location,
        assetnum
    )

@mcp.tool()
def update_work_order(
    wonum: str,
    fields: dict
):
    return maximo_service.update_work_order(
        wonum,
        fields
    )

@mcp.tool()
def change_work_order_status(
    wonum: str,
    status: str
):
    return maximo_service.change_work_order_status(
        wonum,
        status
    )

@mcp.tool()
def add_work_log(
    wonum: str,
    text: str
):
    return maximo_service.add_work_log(
        wonum,
        text
    )







if __name__ == "__main__":
    mcp.run()