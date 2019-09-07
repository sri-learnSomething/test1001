import requests
import pytest
import json
import jsonpath
import openpyxl
from Data_driven import Library

base_url = "http://thetestingworldapi.com/"


def test_add_multiple_students():
    # API access
    api_url = base_url + "api/studentsDetails"
    file = open("C:\\Users\\SridharRaju\\Desktop\\API_tasks\\add_newStudent.json.txt", "r")
    json_request = json.loads(file.read())

    obj = Library.Common("C:\\Users\\SridharRaju\\Desktop\\API_tasks\\newStudentDetail.xlsx", 'Sheet1')
    col = obj.fetch_col_count()
    row = obj.fetch_row_count()
    keyList = obj.fetch_key_names()

    for i in range(2, row + 1):
        updated_json_request = obj.update_request_with_data(i, json_request, keyList)
        response = requests.post(api_url, updated_json_request)
        print(response)