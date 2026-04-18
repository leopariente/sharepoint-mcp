from tools import mcp, warmup

warmup()

if __name__ == "__main__":
    # Endpoint: http://localhost:8000/mcp
    # host="0.0.0.0" accepts all interfaces (required for ngrok).
    mcp.run(transport="streamable-http")
