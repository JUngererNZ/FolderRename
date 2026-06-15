import requests

API_KEY = "0255ef2cb461087caad4c31fa4b1a762ff98f2d9a8babb7701d2a5ca5a2de6d1"
mmsi = 440114000

url = f"https://api.vesselapi.com/v1/vessel/{mmsi}?filter.idType=mmsi"
headers = {"Authorization": f"Bearer {API_KEY}"}

response = requests.get(url, headers=headers)
print(f"Status Code: {response.status_code}")
print(response.json())