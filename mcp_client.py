from mcp import ClientSession
from mcp.client.stdio import stdio_client


class MaximoMCPClient:

    async def call_tool(
        self,
        tool_name: str,
        arguments: dict
    ):
        async with stdio_client(
            command="python",
            args=["mcp_server.py"]
        ) as (reader, writer):

            async with ClientSession(
                reader,
                writer
            ) as session:

                await session.initialize()

                result = await session.call_tool(
                    tool_name,
                    arguments
                )

                return result

    async def list_tools(self):
        async with stdio_client(
            command="python",
            args=["mcp_server.py"]
        ) as (reader, writer):

            async with ClientSession(
                reader,
                writer
            ) as session:

                await session.initialize()

                return await session.list_tools()