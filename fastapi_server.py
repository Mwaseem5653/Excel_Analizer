from fastapi import FastAPI, WebSocket
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import httpx
import os
import asyncio # Add asyncio import
import uvicorn # Add uvicorn import
import requests
from dotenv import load_dotenv

load_dotenv()

app = FastAPI()

# Allow frontend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def home():
    return {"message": "FastAPI server is running on Render!"}

@app.get("/lookup")
def lookup(number: str, code: str = "92"):  # default 92
    rapid_api_key = os.getenv("RAPID_API_KEY", "76590bef8fmshe2ed5dc341f89e2p1d75d3jsn4edec491e6c9")
    url = "https://eyecon.p.rapidapi.com/api/v1/search"

    headers = {
        "x-rapidapi-key": rapid_api_key,
        "x-rapidapi-host": "eyecon.p.rapidapi.com"
    }

    params = {
        "code": code,
        "number": number
    }

    try:
        response = requests.get(url, headers=headers, params=params, timeout=10)
        data = response.json()

        if not data.get("status"):
            return {"status": False, "message": "No record found"}

        return data
    except Exception as e:
        return {"status": False, "message": str(e)}

@app.post("/get-info/")
async def get_phone_info(phone_number: str):
    # This new website has a direct API endpoint we can call
    api_url = "https://simdataupdates.com/wp-admin/admin-ajax.php"
    params = {
        'action': 'fetch_sim_data',
        'term': phone_number
    }
    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Referer': 'https://simdataupdates.com/',
        'X-Requested-With': 'XMLHttpRequest'
    }
    
    try:
        # The request is a simple GET request
        
        response = requests.get(
            api_url,
            params=params,
            headers=headers,
            timeout=10
        )

        # 1️⃣ HTTP check
        if response.status_code != 200:
            return {"error": f"HTTP Error {response.status_code}"}

        # 2️⃣ Empty response check
        if not response.text or len(response.text.strip()) == 0:
            return {"error": "Empty response from website (blocked)"}

        # 3️⃣ Content-Type check
        content_type = response.headers.get("content-type", "")
        if "application/json" not in content_type:
            return {
                "error": "Non-JSON response received (likely blocked)",
                "preview": response.text[:200]
            }

        # 4️⃣ Now safe to parse JSON
        try:
            data = response.json()
        except ValueError:
            return {"error": "Invalid JSON response"}

            
        
        # Check if the API call was successful and if any data was returned
        if data.get("success") and data.get("data") and len(data["data"]) > 0:
            # We will return all records found
            return data["data"]
        else:
            return {"error": "No record found for this number."}

    except requests.exceptions.RequestException as e:
        return {"error": f"Failed to retrieve data from the website: {e}"}
    except ValueError:
        # Catches JSON decoding errors if the response is not valid JSON
        return {"error": "Failed to decode the response from the server. The API might be down or has changed."}

@app.post("/get-vehicle-info/")
async def get_vehicle_info(reg_no: str, category: str):
    api_url = "https://api.mahisite.xyz/sindh/api.php"
    params = {
        'reg_no': reg_no,
        'category': category
    }

    try:
        response = requests.get(api_url, params=params, timeout=15)
        response.raise_for_status()
        
        data = response.json()

        if data.get("statusCode") == 0 and data.get("data") and len(data["data"]) > 0:
            info = data["data"][0]
            return {
                "registrationNumber": info.get("registrationNumber"),
                "ownerName": info.get("ownerName"),
                "ownerCNIC": info.get("ownerCNIC"),
                "ownerAddress": info.get("ownerAddress"),
                "registrationDate": info.get("registrationDate"),
                "engineNumber": info.get("engineNumber"),
                "chassisNumber": info.get("chassisNumber"),
                "branchName": info.get("branchName"),
                "districtName": info.get("districtName"),
                "modelYear": info.get("modelYear"),
                "manufacturerName": info.get("manufacturerName"),
                "modelName": info.get("modelName"),
                "color": info.get("color"),
                "cplcStatus": info.get("cplcStatus"),
            }
        else:
            return {"error": "No vehicle record found."}

    except requests.exceptions.RequestException as e:
        return {"error": f"Failed to retrieve data from the vehicle API: {e}"}
    except ValueError:
        return {"error": "Failed to decode the response from the vehicle API server."}

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