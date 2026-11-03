import requests
import json

swagger_url = "https://data-nsr.udir.no/swagger/v4/swagger.json"

response = requests.get(swagger_url, timeout=60)
response.raise_for_status()

swagger = response.json()

with open("nsr_swagger_v4.json", "w", encoding="utf-8") as f:
    json.dump(swagger, f, ensure_ascii=False, indent=2)

print("API:", swagger.get("info", {}).get("title"))
print("Versjon:", swagger.get("info", {}).get("version"))

print("\nTilgjengelige endepunkter:")
for path, methods in swagger["paths"].items():
    print(path)
    for method in methods:
        if method.upper() in {"GET", "POST", "PUT", "DELETE"}:
            print("   ", method.upper())