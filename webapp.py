import requests
from flask import Flask, render_template, request, redirect
#from microsoftgraph.client import Client

app = Flask(__name__)

client_id = 'acd4b6da-7327-4d6e-b945-47d0197c18a1'
client_secret = 'OxM8Q~k-bPfPxAgESwiNAry05OWth9YqPZ_VfbzH'
tenant_id = '48cc8b58-56d3-4cb6-b8bc-5a76191911e6'
redirect_uri = 'https://shubhankar.fyi'

authorization_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize'
token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'

@app.route('/')
def index():
    # Check if an access token is stored in the session
    access_token = request.cookies.get('access_token')
    if not access_token:
        # Redirect the user to the Microsoft login page
        params = {
            'client_id': client_id,
            'response_type': 'code',
            'redirect_uri': redirect_uri,
            'response_mode': 'query',
            'scope': 'https://graph.microsoft.com/.default',
            'state': '12345'
        }
        url = requests.Request('GET', authorization_url, params=params).prepare().url
        return redirect(url)
    else:
        # Use the access token to create a Microsoft Graph client
        graph_client = Client(client_id, client_secret, access_token=access_token)

        # Get unread emails from Outlook
        emails = graph_client.get('/me/mailFolders/Inbox/messages?$filter=isRead eq false')
        unread_emails = emails['value']

        # Get upcoming events from Outlook Calendar
        events = graph_client.get('/me/events?$filter=start/dateTime ge \'2022-03-30\'')
        upcoming_events = events['value']

        # Get HR announcements
        hr_announcements_url = 'https://shubhankar.fyi'
        hr_announcements = requests.get(hr_announcements_url).json()

        # Get company news
        company_news_url = 'https://shubhankar.fyi'
        company_news = requests.get(company_news_url).json()

        return render_template('dashboard.html', unread_emails=unread_emails, upcoming_events=upcoming_events, hr_announcements=hr_announcements, company_news=company_news)

@app.route('/callback')
def callback():
    # Get the authorization code from the query string
    code = request.args.get('code')

    # Exchange the authorization code for an access token
    data = {
        'client_id': client_id,
        'scope': 'https://graph.microsoft.com/.default',
        'code': code,
        'redirect_uri': redirect_uri,
        'grant_type': 'authorization_code',
        'client_secret': client_secret
    }
    response = requests.post(token_url, data=data)
    response_json = response.json()
    access_token = response_json['access_token']

    # Store the access token in a cookie and redirect to the index page
    response = app.make_response(redirect('/'))
    response.set_cookie('access_token', access_token)
    return response

if __name__ == '__main__':
    app.run()
