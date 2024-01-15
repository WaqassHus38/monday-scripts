import os
from requests import post, Response
from dotenv import load_dotenv

load_dotenv()

apiUrl = os.getenv('MONDAY_URL')
apiKey = os.getenv('MONDAY_API_KEY')
headers = {"Authorization": apiKey, "API-Version": "2023-10"}

def graphql_query(query_data : str) -> Response.json:
    query = {'query' : query_data}
    return post(url=apiUrl, json=query, headers=headers).json()


def graphql_mutation(mutation_data : str) -> Response.json:
    mutation = {'mutation' : mutation_data}
    return post(url=apiUrl, json=mutation, headers=headers).json()




