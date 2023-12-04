import requests

response = requests.post('https://auth.monday.com/oauth2/token?code=7059c000cca5b610e3bc36efbf0b69ea&client_id=9f814ecea11d6f5342d9b8ec1d380ab1&client_secret=861581f9e2b0c505a66e00ab755d324f')

print(response.json())