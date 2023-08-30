"""
This module provides functions for handling Google OAuth2 authorization.

It includes functions to load credentials from a file, refresh credentials,
and authorize with Google using OAuth.
"""
import os

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from json import JSONDecodeError

from oauthlib.oauth2 import AccessDeniedError

from .settings import TOKEN_FILE_PATH, TOKEN_FILE_NAME, \
    CREDENTIALS_FILE_PATH, CREDENTIALS_FILE_NAME, SCOPES


def load_credentials_from_file(file_path: str) -> Credentials | None:
    """ Load credentials from a file. """
    if os.path.exists(file_path):
        return Credentials.from_authorized_user_file(file_path, SCOPES)
    return None


def refresh_credentials(creds: Credentials) -> Credentials:
    """ Refresh the credentials. """
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return creds


def authorize_with_oauth() -> Credentials:
    """
    Handle the process of authorizing with Google using OAuth 2.0.

    Check if valid credentials exist and refresh them if necessary. If
    credentials are not available or valid, it initiates the
    authorization flow, prompting the user to grant access.
    Once authorized, the credentials are saved as a token for future use.
    """
    token_path = os.path.join(TOKEN_FILE_PATH, TOKEN_FILE_NAME)
    credentials_path = os.path.join(CREDENTIALS_FILE_PATH,
                                    CREDENTIALS_FILE_NAME
                                    )

    creds = load_credentials_from_file(token_path)
    if not creds or not creds.valid:
        creds = refresh_credentials(creds) if creds else None

        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path,
                SCOPES
            )
            creds = flow.run_local_server(port=0)

            with open(token_path, 'w', encoding='utf-8') as token:
                token.write(creds.to_json())

    return creds


def get_credentials() -> Credentials:
    """ Get authorized Google credentials. """
    credentials = None

    try:
        credentials = authorize_with_oauth()
        if not credentials:
            print("Failed to authorize")

    except JSONDecodeError:
        print('Token is invalid or expired, recreating a new one...')
        os.remove(os.path.join(TOKEN_FILE_PATH, TOKEN_FILE_NAME))
    except KeyboardInterrupt:
        print("Timeout while trying to authorize, retrying...")
        credentials = get_credentials()
    except AccessDeniedError:
        print("You denied authorization or access blocked. Try again.")
    except FileNotFoundError:
        print('Credentials not found. Check the path and try again.')

    return credentials
