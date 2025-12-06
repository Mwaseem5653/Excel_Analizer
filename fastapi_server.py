from fastapi import FastAPI, WebSocket
from fastapi.responses import JSONResponse
import httpx
import os
import asyncio # Add asyncio import
import uvicorn # Add uvicorn import

app = FastAPI()

@app.get("/")
async def home():
    return {"message": "FastAPI server is running on Render!"}

@app.get("/proxy")
async def proxy(accountNumber: str):
    api_url = "https://easyload.com.pk/dingconnect.php"
    params = {"action": "GetProviders", "accountNumber": accountNumber}
    
    async with httpx.AsyncClient() as client:
        try:
            response = await client.get(api_url, params=params)
            response.raise_for_status()
            return JSONResponse(content=response.json())
        except httpx.HTTPStatusError as e:
            return JSONResponse(
                content={"error": f"HTTP error occurred: {e.response.status_code}"},
                status_code=e.response.status_code
            )
        except httpx.RequestError as e:
            return JSONResponse(
                content={"error": f"An error occurred while requesting {e.request.url!r}."},
                status_code=500
            )

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await websocket.accept()
    while True:
        data = await websocket.receive_text()
        await websocket.send_text(f"Message text was: {data}")

if __name__ == "__main__":
    # Render provides PORT environment variable
    # We will let the OS choose a free port by setting port to 0
    configured_port = int(os.environ.get("PORT", 0))

    config = uvicorn.Config(app, host="0.0.0.0", port=configured_port, log_level="info")
    server = uvicorn.Server(config)

    # To get the actual port, we need to run the server programmatically
    # and wait for it to start.
    async def start_server_and_write_port():
        # Start the server in a separate task
        server_task = asyncio.create_task(server.serve())

        # Wait for the server to be ready and its sockets to be bound
        while not server.started:
            await asyncio.sleep(0.1)
        
        # The actual port is available after startup
        actual_port = server.servers[0].sockets[0].getsockname()[1]
        print(f"Uvicorn server started on port: {actual_port}")

        # Write the actual port to a file
        with open(".uvicorn_port", "w") as f:
            f.write(str(actual_port))

        await server_task # Keep the server running

    asyncio.run(start_server_and_write_port())
