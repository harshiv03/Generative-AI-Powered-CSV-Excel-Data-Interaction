!pip install openai
!pip install pandas
!pip install langchain
!pip install langchain_experimental
!pip install langchain_openai
!pip install streamlit
!pip install openpyxl
!pip install python-dotenv
!pip install openai langchain

import pandas as pd
from langchain_experimental.agents import create_csv_agent
from langchain_openai import AzureOpenAI
import os
import streamlit as st
import openpyxl
import io

def main():
    api_key = "add your api-key"
    azure_endpoint = "add your endpoint"
    api_version = "add your api-version"
    return api_key, azure_endpoint, api_version
api_key, azure_endpoint, api_version = main()

file_path = "/content/product and customer report.xlsx"

if file_path.endswith('.csv'):
    csv_files = [file_path]
elif file_path.endswith('.xlsx'):
    # Read all sheets from the Excel file
    df_dict = pd.read_excel(file_path, sheet_name=None)
    csv_files = []
    for sheet_name, df in df_dict.items():
        print(f"Data from sheet: {sheet_name}")
        print(df)
        # Save the dataframe to a csv file
        csv_file = f"{sheet_name}.csv"
        df.to_csv(csv_file, index=False)
        csv_files.append(csv_file)
else:
    print("Unsupported file type")


agent = create_csv_agent(
    AzureOpenAI(
        model="add your model",
        deployment_name="add your model deployment-name",
        api_key=api_key,
        azure_endpoint=azure_endpoint,
        api_version=api_version,
        temperature=1
    ),
    csv_files,
    verbose=True,
    handle_parsing_errors=True,  # Handle parsing errors gracefully
    allow_dangerous_code=True,
    return_intermediate_steps=True  # Return intermediate steps
)

def ask_question():
    user_question = input("Enter your question: ")
    if user_question:
        result = agent({"input": user_question})  # Pass input as a dictionary
        print("Answer:")
        if 'output' in result:
            print(result['output'])  # Access the final answer
        else:
            print("No answer found in the result.")
       

# Run the question-answering loop
while True:
    ask_question()
    response = input("Do you want to ask another question? (yes/no): ")
    if response.lower()!= "yes":
        break