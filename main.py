import os
import pickle
import base64
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from tqdm import tqdm
import random
import re
from bs4 import BeautifulSoup
from selenium import webdriver
from PyInquirer import prompt, Separator
import webbrowser


# Define the scopes for the Gmail API
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

# Connect to the Gmail API
def connect_to_gmail():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('gmail', 'v1', credentials=creds)
    return service

# Retrieve inbound emails
def get_inbound_emails(service, lim = None):
    query = "in:inbox to:me"
    results = service.users().messages().list(userId='me', q=query, maxResults=1000).execute()
    messages = results.get('messages', [])
    all_emails = []
    while messages:
        if lim and len(all_emails) < lim:
            break
        pbar = tqdm(messages, desc="Retrieving inbound emails", unit="message")
        for msg in pbar:
            pbar.set_postfix({"emails found": len(all_emails)})
            txt = service.users().messages().get(userId='me', id=msg['id']).execute()
            if get_unsubscribe_url(txt):
                all_emails.append(txt)
                if lim and len(all_emails) >= lim:
                    break
        if 'nextPageToken' in results:
            page_token = results['nextPageToken']
            results = service.users().messages().list(userId='me', q=query, maxResults=1000, pageToken=page_token).execute()
            messages = results.get('messages', [])
        else:
            break
    return all_emails


# Retrieve all emails
def get_all_emails(service, lim = None):
    results = service.users().messages().list(userId='me', maxResults=1000).execute()
    messages = results.get('messages', [])
    if lim:
        messages = random.sample(messages, lim)
    all_emails = []
    while messages and len(all_emails) < lim:
        for msg in tqdm(messages):
            txt = service.users().messages().get(userId='me', id=msg['id']).execute()
            all_emails.append(txt)
        if 'nextPageToken' in results:
            page_token = results['nextPageToken']
            results = service.users().messages().list(userId='me', maxResults=1000, pageToken=page_token).execute()
            messages = results.get('messages', [])
        else:
            break
    return all_emails

# # Check if an email is inbound
# def is_inbound_email(email):
#     sender = get_header(email, "from")
#     reply_to = get_header(email, "reply-to")
#     sender_name = get_header(email, "sender/name")
#     sender_email = get_header(email, "sender/email")
#     return sender != "me" and sender_email != "me" and sender_name != "me" and reply_to != "me"

# # Get the value of a header field
def get_header(email, name):
    headers = email["payload"]["headers"]
    for header in headers:
        if header["name"].lower() == name.lower():
            return header["value"]
    return ""

# Group emails by domain
def group_by_domain(emails):
    domain_dict = {}
    for email in emails:
        sender = get_header(email, "FROM")
        domain = sender.split('@')[1]
        if domain in domain_dict:
            domain_dict[domain].append(email)
        else:
            domain_dict[domain] = [email]
    return domain_dict

# Sort by total number of messages
def sort_by_total_messages(domain_dict):
    sorted_domains = sorted(domain_dict.items(), key=lambda x: len(x[1]), reverse=True)
    return sorted_domains

# Get the unsubscribe URL from an email
def get_unsubscribe_url(email):
    # Check if email has an HTML part
    parts = email["payload"].get("parts")
    if parts:
        for part in parts:
            if part["mimeType"] == "text/html":
                # Parse HTML content
                html = part["body"]["data"]
                html = base64.urlsafe_b64decode(html).decode('utf-8')
                soup = BeautifulSoup(html, 'html.parser')
                # Find the first link that contains "unsubscribe" in the href attribute
                for link in soup.find_all('a'):
                    if 'unsubscribe' in link.get('href', '').lower():
                        return link.get('href')
    # If email does not have an HTML part or no unsubscribe link is found, return None
    return None

# # Unsubscribe from mailing list
# def unsubscribe_from_mailing_list(service, domain):
#     search_query = f"from:{domain}"
#     result = service.users().messages().list(q=search_query, userId='me').execute()
#     messages = result.get('messages', [])
#     if not messages:
#         print("No messages found from this domain.")
#         return
#     print(f"{len(messages)} messages found from {domain}.")
#     confirm = input("Are you sure you want to unsubscribe from this mailing list? (y/n) ")
#     if confirm.lower() == "y":
#         for message in messages:
#             message_id = message['id']
#             labels = {'removeLabelIds': ['INBOX'], 'addLabelIds': ['UNSUBSCRIBE']}
#             service.users().messages().modify(userId='me', id=message_id, body=labels).execute()
#         print(f"Unsubscribed from {domain}.")
#     else:
#         print("Aborted.")

def user_form(domains_obj):
    domains = [el[0] for el in domains_obj]
    emails = [el[1] for el in domains_obj]
    # Define the prompt questions
    questions = [
        {
            'type': 'checkbox',
            'message': 'Select domains to unsubscribe from:',
            'name': 'domains',
            'choices': [{'name': f"{domain}"} for i, domain in enumerate(domains)]
        }
    ]

    # Prompt the user to select domains
    answers = prompt(questions)

    # Read in the selected domains
    selected_domains = answers['domains']
    output = [el for el in domains_obj if el[0] in selected_domains]
    return output

def extract_unsub_links(domains_to_unsub):
    unsub_links = set()
    for domain, emails in domains_to_unsub:
        for email in emails:
            unsub_links.add(get_unsubscribe_url(email))
    return unsub_links

# Archive an email
def archive_email(service, email_id):
    message = service.users().messages().modify(
        userId='me',
        id=email_id,
        body={
            'removeLabelIds': ['INBOX']
        }).execute()

# Ask user if they want to archive emails for domains they've unsubscribed from
def ask_to_archive(service, unsubscribe_domains):
    print("You have unsubscribed from the following domains:")
    print([el[0] for el in unsubscribe_domains])
    answer = input("Do you want to archive all emails for these domains? (y/n)")
    if answer.lower() == 'y':
        for domain, emails in tqdm(unsubscribe_domains):
            for email in emails:
                archive_email(service, email['id'])
        print("All emails for the unsubscribed domains have been archived.")
    else:
        print("No emails have been archived.")


# Main function
def main():
    print("Connecting to gmail...")
    service = connect_to_gmail()
    print("Getting all emails...")
    all_emails = get_inbound_emails(service)
    print(f"All emails: {len(all_emails)}")
    print("Grouping emails by domain...")
    domain_dict = group_by_domain(all_emails)
    sorted_domains = sort_by_total_messages(domain_dict)
    domains_to_unsub = user_form(sorted_domains)
    unsub_links = extract_unsub_links(domains_to_unsub)

    for i, url in enumerate(unsub_links):
        webbrowser.open(url)

    ask_to_archive(service, domains_to_unsub)


if __name__ == '__main__':
    main()