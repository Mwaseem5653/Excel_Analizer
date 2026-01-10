import requests
import httpx
import os

def get_phone_info(phone_number: str):
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
        response = requests.get(
            api_url,
            params=params,
            headers=headers,
            timeout=10
        )

        if response.status_code != 200:
            return {"error": f"HTTP Error {response.status_code}"}

        if not response.text or len(response.text.strip()) == 0:
            return {"error": "Empty response from website (blocked)"}

        content_type = response.headers.get("content-type", "")
        if "application/json" not in content_type:
            return {
                "error": "Non-JSON response received (likely blocked)",
                "preview": response.text[:200]
            }

        data = response.json()
        
        if data.get("success") and data.get("data") and len(data["data"]) > 0:
            return data["data"]
        else:
            return {"error": "No record found for this number."}

    except requests.exceptions.RequestException as e:
        return {"error": f"Failed to retrieve data: {e}"}
    except ValueError:
        return {"error": "Failed to decode the response."}

def get_vehicle_info(reg_no: str, category: str):
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
        return {"error": f"Failed to retrieve vehicle data: {e}"}
    except ValueError:
        return {"error": "Failed to decode response."}

async def get_providers_proxy(accountNumber: str):
    api_url = "https://easyload.com.pk/dingconnect.php"
    params = {"action": "GetProviders", "accountNumber": accountNumber}
    
    async with httpx.AsyncClient() as client:
        try:
            response = await client.get(api_url, params=params)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            return {"error": str(e)}

def get_eyecon_info(number: str, code: str = "92"):
    rapid_api_key = os.getenv("RAPID_API_KEY")
    if not rapid_api_key:
        return {"error": "RAPID_API_KEY not found in environment variables."}

    url = "https://eyecon3.p.rapidapi.com/api/v1/search"
    headers = {
        "x-rapidapi-key": rapid_api_key,
        "x-rapidapi-host": "eyecon3.p.rapidapi.com"
    }
    params = {
        "code": code,
        "number": number
    }

    try:
        response = requests.get(url, headers=headers, params=params, timeout=10)
        data = response.json()
        print(f"Eyecon API Response for {number}: {data}")
        
        # Return data even if status is false/missing so we can debug the error (e.g., Auth failure, Quota)
        if not data.get("status"):
             # If no error key, maybe add one for consistency, but for now just return raw data
             pass
            
        return data
    except Exception as e:
        return {"error": str(e)}
