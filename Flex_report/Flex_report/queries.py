from graphql import graphql_query
import time

def items_pagination(data):
    return (
        data["data"]["boards"][0]["groups"][0]["items_page"]["items"] is not None
        and len(data["data"]["boards"][0]["groups"][0]["items_page"]["items"]) == 0
    )

def boards_pagination(data):
    return (
        data["data"]["boards"] is not None
        and len(data["data"]["boards"]) == 0
    )

def paginated_query(query: str, pagination_function : callable):
    limit = 10
    page = 1

    array = []

    while True:
        _query = query.replace("##limit##", str(limit)).replace("##page##", str(page))
        data = graphql_query(_query)
        if pagination_function(data):
            break
        array.append(data)
        page += 1

    return array


def get_item_query(board_id: str, group_id: str):
    query = """
        query {
        boards (ids : [%s]) {
            groups (ids : ["%s"]) {
                items_page (limit: ##limit##) {
                    cursor
                    items {
                        id
                        name
                        column_values {
                        id 
                        text
                        value
                        column {
                            title
                        }
                    }
                    }
                }
            }
        }
    }
    """ % (
        board_id,
        group_id,
    )
    return query

def get_item_by_board_id_query(board_id: str):
    query = """
        query {
            boards (ids : [%s]) {
                
                items (limit: ##limit##, page: ##page##) {
                    id
                    name
                    column_values {
                        id 
                        text
                        value
                        title
                    }
                }
                
            }
        }
    """ % (
        board_id
    )
    return query


def get_boards_graphql():
    query = """
        query {
            boards (workspace_ids: 172030, limit: ##limit##, page: ##page##) {
                id
                name
                groups {
                    id
                    title
                }
            }
        }
    """
    return query



def parse_board_data(data) -> list:
    boards = []
    for x in data:
        boards.extend(x["data"]["boards"])
    return boards



def parse_item_data(data) -> list:
    items = []
    for x in data:
        if len(x["data"]["boards"]) > 0:
            items.extend(x["data"]["boards"][0]["groups"][0]['items_page']["items"])
    return items


def change_item_value(board_id: str, item_id: str, column_id: str, value: str):
    mutation = """
        mutation {
            change_simple_column_value (board_id: %s, item_id: %s, column_id: "%s", value: "%s") {
                id
            }
        }
    """ % (
        board_id,
        item_id,
        column_id,
        value,
    )
    return mutation


def move_item_to_group(item_id: str, group_id: str):
    mutation = """
        mutation {
            move_item_to_group (item_id: %s, group_id: "%s") {
                id
            }
        }
    """ % (
        item_id,
        group_id,
    )
    return mutation


def create_group(board_id: str, group_name: str):
    mutation = """
        mutation {
            create_group (board_id: %s, group_name: "%s", position_relative_method: before_at) {
                id
            }
        }
    """ % (
        board_id,
        group_name,
    )
    return mutation


def new_paginated_query(query: str, pagination_function: callable , get_cursor : callable = None):
    limit = 5
    page = 1

    array = []

    cursor = ""

    while True:
        
        _query = query.replace("##limit##", str(limit)).replace("##page##", str(page)).replace("##cursor##", str(cursor))
        data = graphql_query(_query)

        if "error_code" in data:
            if data["error_code"] == "ComplexityException":
                print("Complexity budget exceeded, reducing limit and retrying")
                time.sleep(60)
                continue

            elif data["error_code"] == "UserUnauthroizedException":
                return []
            
        if get_cursor is not None:
            if get_cursor(data) is None:
                array.append(data)
                break
            else:
                cursor_string = get_cursor(data)
                cursor = f""", cursor: "{cursor_string}" """

        if pagination_function(data):
            break
        array.append(data)
        page += 1
    
    return array

def new_get_item_query(board_id: str, group_id: str):
    query = f"""
    query {{
      boards(ids: [{board_id}]) {{
        groups(ids: ["{group_id}"]) {{
          items_page(limit: ##limit## ##cursor##) {{
              cursor
              items {{
                  id
                  name
                  created_at
                  column_values {{
                      id
                      text
                      value
                      column {{
                          id
                          title
                      }}
                  }}
              }}
          }}
        }}
      }}
    }}
    """
    return query

def new_item_pagination(data):
    return (
            data['data']['boards'] is not None
            and len(data['data']['boards']) == 0
    )

def get_item_cursor(data):
    return data['data']['boards'][0]['groups'][0]['items_page']['cursor'] if len(data['data']['boards']) > 0 else None


