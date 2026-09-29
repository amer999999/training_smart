import json
import requests

url = 'http://localhost:8017'
db = 'real_estate_db'
username = 'admin'
password = 'admin'

# Step 1: Authenticate
auth_payload = {
    "jsonrpc": "2.0",
    "method": "call",
    "params": {
        "service": "common",
        "method": "authenticate",
        "args": [db, username, password, {}]
    },
    "id": 1
}

response = requests.post(f'{url}/jsonrpc', json=auth_payload)
uid = response.json()['result']
print('uid:', uid)
# Step 2: Create a maintenance request (tenant mobile app use case)
create_payload = {
    "jsonrpc": "2.0",
    "method": "call",
    "params": {
        "service": "object",
        "method": "execute_kw",
        "args": [
            db, uid, password,
            'maintenance.request',
            'create',
            [{
                'lease_id': 31,  # Use a valid property ID
                'issue_type': 'air_condition',
                'description': 'Air conditioning stopped working',
                'urgency': 'high'
            }]
        ]
    },
    "id": 2
}

response = requests.post(f'{url}/jsonrpc', json=create_payload)
result_json = response.json()
print(result_json)
if 'error' in result_json:
    print("Error creating maintenance request:")
    print(json.dumps(result_json['error'], indent=4))
else:
    new_request_id = result_json['result']
    print(f"Successfully created maintenance request with ID: {new_request_id}")