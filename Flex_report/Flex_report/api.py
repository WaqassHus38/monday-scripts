import os
import requests
import logging
from math import floor
import sys
from enum import Enum
from requests.auth import HTTPBasicAuth

# Set up basic logger
logger = logging.getLogger('theflexliving.keynest')

# Setup stdout logger
soh = logging.StreamHandler(sys.stdout)
# Can optionally set logging levels per handler
# soh.setLevel(logging.WARN)
logger.addHandler(soh)

# Get log level from env vars
log_level = os.environ.get('LOG_LEVEL', 'INFO').upper()
if os.environ.get('DEBUG'):
    if log_level:
        logger.warn("Overriding LOG_LEVEL setting with DEBUG")
    log_level = 'DEBUG'

try:
    logger.setLevel(log_level)
except ValueError:
    logger.setLevel(logging.INFO)
    logger.warn("Variable LOG_LEVEL not valid - Setting Log Level to INFO")


# Set URL for the keynest
URL = 'https://api.keynest.com/api/v3/'


class APIEndPoint():
    def __init__(self, ApiKey):

        # Setup host
        self.url = URL
        # API Key for keynest

        # Setup Session object for all future API calls
        self.session = requests.Session()
        self.session.headers.update({'ApiKey': ApiKey})

    def _make_request(self, endpoint, method, query_params=None, body=None):
        '''Handles all requests to Example API
        '''
        url = self.url + endpoint
        req = requests.Request(method, url, params=query_params, json=body)
        prepped = self.session.prepare_request(req)

        # Log request prior to sending
        self._pprint_request(prepped)

        # Actually make request to endpoint
        r = self.session.send(prepped)

        # Log response immediately upon return
        self._pprint_response(r)

        # Handle all response codes as elegantly as needed in a single spot
        if r.status_code == requests.codes.ok:
            try:
                resp_json = r.json()
                logger.debug('Response: {}'.format(resp_json))
                return resp_json
            except ValueError:
                return r.text

        elif r.status_code == 401:
            logger.info("Authentication Unsuccessful!")
            try:
                resp_json = r.json()
                logger.debug('Details: ' + str(resp_json))
                raise AuthenticationError(resp_json)
            except ValueError:
                raise

        # TODO handle rate limiting gracefully

        # Raises HTTP error if status_code is 4XX or 5XX
        elif r.status_code >= 400:
            logger.error('Received a ' + str(r.status_code) + ' error!')
            try:
                logger.debug('Details: ' + str(r.json()))
            except ValueError:
                pass
            r.raise_for_status()

    def _pprint_request(self, prepped):
        '''
        method endpoint HTTP/version
        Host: host
        header_key: header_value
        body
        '''
        method = prepped.method
        url = prepped.path_url
        # TODO retrieve HTTP version
        headers = '\n'.join('{}: {}'.format(k, v) for k, v in
                            prepped.headers.items())
        # Print body if present or empty string if not
        body = prepped.body or ""

        logger.info("Requesting {} to {}".format(method, url))

        logger.debug(
            '{}\n{} {} HTTP/1.1\n{}\n\n{}'.format(
                '-----------REQUEST-----------',
                method,
                url,
                headers,
                body
            )
        )

    def _pprint_response(self, r):
        '''
        HTTP/version status_code status_text
        header_key: header_value
        body
        '''
        # Not using requests_toolbelt.dump because I want to be able to
        # print the request before submitting and response after
        # ref: https://stackoverflow.com/a/35392830/8418673

        httpv0, httpv1 = list(str(r.raw.version))
        httpv = 'HTTP/{}.{}'.format(httpv0, httpv1)
        status_code = r.status_code
        status_text = r.reason
        headers = '\n'.join('{}: {}'.format(k, v) for k, v in
                            r.headers.items())
        body = r.text or ""
        # Convert timedelta to milliseconds
        elapsed = floor(r.elapsed.total_seconds() * 1000)

        logger.info(
            "Response {} {} received in {}ms".format(
                status_code,
                status_text,
                elapsed
            )
        )

        logger.debug(
            '{}\n{} {} {}\n{}\n\n{}'.format(
                '-----------RESPONSE-----------',
                httpv,
                status_code,
                status_text,
                headers,
                body
            )
        )

    def get_key_nests(self, endpoint, method="GET", query_params=None, body=None):
        '''
        get a list of keynests
        '''
        return self._make_request(endpoint, method, query_params, body)


class AuthenticationError(Exception):
    pass
