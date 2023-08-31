import json
import pathlib


def import_wf_file(fileName,fileDir):
    file_path = pathlib.Path.joinpath(pathlib.Path(fileDir), pathlib.Path(fileName))
    with open(file_path, "r") as file:
        data = json.load(file)
    return data


def save_wf_file(fileName,fileDir, data):
    filepath = pathlib.Path.joinpath(pathlib.Path(fileDir), pathlib.Path(fileName))
    if filepath.exists():
        with open(filepath, "w") as file:
            json.dump(data, file, indent=4)

BASE_TEMPLATE = """
{
    "cells":[
    ]
}
"""

# Currently It is not working

# BASE_TEMPLATE = """
# {
#     "cells":[
#         {
#             "type": "markdown",
#             "content": "This is a markdown cell and below will be a task cell."
#         },
#         {
#             "type": "task",
#             "task_type": "write",
#             "task_data":{
#                 "interval": 0.2,
#                 "text": "Hello World!"
#             }
#         }
#     ]
# }
# """


def create_wf_file(fileName,fileDir):
    save_wf_file(fileName, fileDir, json.loads(BASE_TEMPLATE))


def process_data(data):
    print("Processing Done!")