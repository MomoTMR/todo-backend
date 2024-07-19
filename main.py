import requests

url = "https://aisoip.adilet.gov.kz/rest/debtor/findErd"
params = {'page': 0, 'size': 10}

try:
    response = requests.get(url, params=params)
    response.raise_for_status()  # Raises error for bad responses (4xx or 5xx)
    data = response.json()
    print(data)
except requests.HTTPError as http_err:
    print(f'HTTP error occurred: {http_err}')  # For 4xx, 5xx errors
except Exception as err:
    print(f'Other error occurred: {err}')  # For other errors