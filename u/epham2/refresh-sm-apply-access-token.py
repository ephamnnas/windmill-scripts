import wmill
import requests

# Load current values from Windmill
client_id = wmill.get_variable("u/epham/SMAPPLY_CLIENT_ID")
client_secret = wmill.get_variable("u/epham/SMAPPLY_CLIENT_SECRET")
refresh_token = wmill.get_variable("u/epham/SMAPPLY_REFRESH_TOKEN")

# SM Apply token endpoint
url = "https://nnas.smapply.io/api/o/token/"
data = {
    "grant_type": "refresh_token",
    "client_id": client_id,
    "client_secret": client_secret,
    "refresh_token": refresh_token
}

def main():
    response = requests.post(url, data=data)

    if response.status_code == 200:
        tokens = response.json()
        access_token = tokens["access_token"]
        new_refresh_token = tokens["refresh_token"]

        # Store updated tokens in Windmill
        wmill.set_variable("u/epham/SMAPPLY_ACCESS_TOKEN", access_token)
        wmill.set_variable("u/epham/SMAPPLY_REFRESH_TOKEN", new_refresh_token)

        print("Access and refresh tokens updated successfully.")
    else:
        print(f"Failed to refresh tokens: {response.status_code} {response.text}")
