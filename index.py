from flask import Flask, Response,  request
import PyPDF2, json
import os, io
import openai
import pandas as pd
import xlsxwriter
from openpyxl import Workbook
import logging

app = Flask(__name__)
openai.api_key = os.getenv("OPENAI_API_KEY")

@app.route('/Ping', methods=(['GET']))
def index():
    return "API is running"

@app.route('/GetPlan', methods=(['POST']))
def GetPlan():
    if request.method == 'POST':
        jsonData = request.get_json()
        modifiedPrompt = GetPromptText(jsonData["goal"], jsonData["time"])
        response = openai.Completion.create(
            model="text-davinci-003",
            prompt=modifiedPrompt,
            temperature=0.7,
            max_tokens=3000
        )

        responseText = response.choices[0].text
        try:
            responseJson = json.loads(responseText)        
        except:
            logging.error("error while parsing GPT response JSON.")

        workbook = xlsxwriter.Workbook("output.xlsx")
        sheet = workbook.add_worksheet()

        process_node(sheet, responseJson, 1, 0)

        workbook.close()

        # prepare the response with the excel file
        with open("output.xlsx", "rb") as excel_file:
            response = Response(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response.headers["Content-Disposition"] = "attachment; filename=data.xlsx"

        return response

def GetPromptText(input, time):

    prompt = None
    SITE_ROOT = os.path.abspath(os.path.dirname(__file__))
    formatPath = os.path.join(SITE_ROOT, "static","prompts", "format.json")
    promptPath = os.path.join(SITE_ROOT, "static","prompts", "task-planner-prompt.txt")

    f = open(formatPath)
    format = json.load(f)

    with open(promptPath, "r") as file:
        prompt = f"{file.read()}"
    
    prompt1 = prompt.format(input = input, time = time, format = json.dumps(format))
    return prompt1

def add_row(sheet, key, description, row, col):
    sheet.write(row, col, key)
    sheet.write(row, col+1, description)

def process_node(sheet, node, row, col):
    row1 = row
    for key, value in node.items():
        if(key != "Description"):
            description = value.get("Description", "")
            add_row(sheet, key, description, row, col)
            if isinstance(value, dict):
                row1 = process_node(sheet, value, row+1, col+2)    
            row+=(row1-row)
            row+=1
    return row
    

