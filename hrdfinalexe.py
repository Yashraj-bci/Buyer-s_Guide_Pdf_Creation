# import files
####### INITIALISATIONS ######
import os
import openai
import json
from docxtpl import DocxTemplate, RichText
from docx import Document
import time
import random
import string
from docx2pdf import convert
import tkinter as tk
import pandas as pd
import pandas as pd
from tkinter import simpledialog


excel_folder_name = "excel_data"
current_folder = os.getcwd()

# Get the current folder path
excel_folder = os.path.join(current_folder, excel_folder_name)

# Get the list of Excel files in the folder
excel_files = [file for file in os.listdir(excel_folder) if file.endswith(".xlsx")]

if excel_files:

    # Fetch the first Excel file in the list

    first_excel_file = excel_files[0]

    # Construct the full file path
    excel_path = os.path.join(excel_folder, first_excel_file)



#read the excel file
df = pd.read_excel(excel_path)
sl = df.values.tolist()


####################### DEFINE ALL FUNCTIONS ####################################################################################

def convert_docx_to_pdf(docx_file_path, pdf_file_path):
    try:
        # Change file permissions
        os.chmod(docx_file_path, 0o777)
        # Convert DOCX to PDF
        convert(docx_file_path, pdf_file_path)
        print(f'Successfully converted {docx_file_path} to {pdf_file_path}')
    except Exception as e:
        print(f'Error converting file: {str(e)}')
def save_docx_from_response(data, docx_folder_path):
    """
    Args:
        data (dictionary): buyers guide response
        docx_folder_path (string): folder to save the docx file
    :return, (string): docx file name
    """

    # Load the DOCX document
    doc = DocxTemplate('HRD_EXE_Template.docx')

    # Define the placeholders and their corresponding values

    placeholders = {
        'title': RichText("Bespoke Buyer Guide", color='#FFFFFF', size=60, bold=True),
        'subtitle': RichText(data["User_submit_info"]["firstName"] + " " + data["User_submit_info"]["lastName"],
                             color='#FFFFFF', bold=True),
        'INTROP1': RichText(data["User_submit_info"]["learningAreas"] + '\n', size=60, bold=True),
        'PARA1': RichText(data["bg_content"]["introduction"]),
        'JobTitle': RichText(data["User_submit_info"]["title"]),
        'CompName': RichText(data['User_submit_info']["company"]),
        'InterestArea': RichText(data['User_submit_info']["learningAreas"]),
        'NoEmployee': data['User_submit_info']['No_of_Employees__c_contact'],
        'CodeCountry': data['User_submit_info']['country'],
        'page_break': RichText('\f')

    }

    Table_keys = {"Product Name": 'P', 'Description': 'D', 'Cost': 'C', 'Integration Capabilities': 'IC',
                  'Suitability Index': 'SI', 'Community Review': 'CR', 'Rating': 'R', 'Scalability': 'SC'}
    table_dic = {}
    for i in range(0, 5):
        table_dic['P' + str(i)] = data["bg_content"]["table_content"][i]["Product Name"]
        table_dic['D' + str(i)] = data["bg_content"]["table_content"][i]["Description"]
        table_dic['C' + str(i)] = data["bg_content"]["table_content"][i]["Cost"]
        table_dic['IC' + str(i)] = data["bg_content"]["table_content"][i]["Integration Capabilities"]
        table_dic['SI' + str(i)] = data["bg_content"]["table_content"][i]["Suitability Index"]
        table_dic['CR' + str(i)] = data["bg_content"]["table_content"][i]["Community Review"]
        table_dic['R' + str(i)] = data["bg_content"]["table_content"][i].get("Rating", "NA")
        table_dic['SC' + str(i)] = data["bg_content"]["table_content"][i]["Scalability"]

    placeholders.update(table_dic)

    compare = {
        'CA': RichText(data["bg_content"]["comparative_analysis"], size=20)
    }

    placeholders.update(compare)

    # Render the placeholders with the provided values
    doc.render(placeholders)



    # Save the modified document inside the folder


    file_name = data["User_submit_info"]["firstName"]+"_"+data["User_submit_info"]["lastName"] + '.docx'
    file_path = docx_folder_path + file_name
    doc.save(file_path)

    return file_name
  
#Inserting keys using tkinter  
   
openai_api_key = None
azure_openai_key = None

def get_api_keys():
    global openai_api_key, azure_openai_key

    if openai_api_key is None or azure_openai_key is None:
        root = tk.Tk()
        root.withdraw()

        # Create a custom dialog with two input fields for API keys
        dialog = tk.Toplevel(root)
        dialog.title("API Keys")

        openai_label = tk.Label(dialog, text="OpenAI API Key:")
        openai_label.pack()

        openai_entry = tk.Entry(dialog)
        openai_entry.pack()

        azure_label = tk.Label(dialog, text="Azure OpenAI API Key:")
        azure_label.pack()

        azure_entry = tk.Entry(dialog)
        azure_entry.pack()

        # Function to save the API keys when the Submit button is clicked
        def save_keys():
            global openai_api_key, azure_openai_key
            openai_api_key = openai_entry.get()
            azure_openai_key = azure_entry.get()
            dialog.destroy()

        submit_button = tk.Button(dialog, text="Submit", command=save_keys)
        submit_button.pack()

        # Wait for the user to input the keys and click Submit
        dialog.wait_window(dialog)

    return openai_api_key, azure_openai_key


def chat_complete(conversation):
    time.sleep(1.5)
    try:
        openai_api_key, azure_openai_key = get_api_keys()
        try:
            openai.api_key = openai_api_key
            chat_engine = "gpt-3.5-turbo-16k"
            openai.api_base = "https://api.openai.com/v1"
            openai.api_type = "open_ai"
            openai.api_version = None
            chat_completion = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-16k",
            messages=conversation,
            top_p=1)
            return chat_completion["choices"][0]['message']['content']

        except Exception as e:
            print("failed with OpenAI"+str(e))
            response = {"Exception": str(e)}

        try:

            openai.api_type = "azure"
            openai.api_key = azure_openai_key
            openai.api_base = "https://bc-api-management-uksouth.azure-api.net"
            openai.api_version = "2023-03-15-preview"
            chat_completion = openai.ChatCompletion.create(engine="gpt-35-turbo", messages=conversation)
            return chat_completion["choices"][0]['message']['content']
        except Exception as e:
            print("failed with Azure OpenAI" + str(e))
            response = {"Exception": str(e)}
    except:  
        return response



def generate_intro(conversation):
    intro_response = conversation[20]["content"]
    prompt = f'''Rewrite the below introduction so that it reads like the introduction of a report. Don't include any title or section header. Just respond with the contents of the introduction.\n\n[start of introduction]\n{intro_response}\n[end of introduction]\n'''
    # chat_completion = openai.ChatCompletion.create(engine=chat_engine, messages=[{"role":"user", "content":prompt}])
    chat_completion = chat_complete([{"role":"user", "content":prompt}])
    return chat_completion

def generate_solutions(conversation):
    solution_response = conversation[22]["content"]
    prompt = f'''Rewrite the below content so that it reads like the content of a solutions section of a report. Don't include any title or section header. Just respond with the contents of this section.\n\n[start of content]\n{solution_response}\n[end of content]\n'''
    # chat_completion = openai.ChatCompletion.create(engine=chat_engine, messages=[{"role":"user", "content":prompt}])
    chat_completion = chat_complete([{"role": "user", "content": prompt}])
    return chat_completion


def generate_json_for_table(conversation):
    print("generate json")
    prompt = f'''Given below is a question and answer discussion. Based on this discussion please give me a Python list of 5 JSON objects. Each JSON object should signify one solution. The list should only include the smaller/medium size providers who are disrupting the market with new technologies or solutions. Each JSON object must have the following keys - 'Product Name', 'Cost', 'Description', 'Integration Capabilities', 'Suitability Index', 'Community Review', 'Scalability' and 'Rating'. Ensure that keys and values are enclosed in double quotes and make sure to include all keys in the JSON response. Additionally, ensure that the cost is real, and each description is differently written. Lastly, the Community Reviews should just be a sentence with the review of the solution and should not include where the review is coming from or any score. Use the below discussion to generate the list of JSON as required:\n\n[start of discussion]\n{conversation[0]['content']}\n\n{conversation[1]['content']}\n\n{conversation[2]['content']}\n\n{conversation[3]['content']}\n\n{conversation[4]['content']}\n\n{conversation[5]['content']}\n\n{conversation[6]['content']}\n\n{conversation[7]['content']}\n\n{conversation[8]['content']}\n\n{conversation[9]['content']}\n\n{conversation[10]['content']}\n\n{conversation[11]['content']}\n\n{conversation[12]['content']}\n\n{conversation[13]['content']}\n\n{conversation[14]['content']}\n\n{conversation[15]['content']}\n\n{conversation[16]['content']}\n[end of discussion]\n\nYour response MUST begin with "[{" and end with "}]"'''
    # chat_completion = openai.ChatCompletion.create(engine=chat_engine, messages=[{"role":"user", "content":prompt}], temperature=0.01)
    chat_completion = chat_complete([{"role": "user", "content": prompt}])
    json_string = chat_completion
    json_string = json_string.replace('\n', '')
    json_object = json.loads(json_string)
    return(json_object)

def generate_comparison(conversation):
    comparison_response = conversation[18]["content"]
    prompt = f'''Rewrite the below content so that it reads like the comparison and summary section of a report. Don't include any title or section header. Just respond with the contents of this section.\n\n[start of content]\n{comparison_response}\n[end of content]\n'''
    # chat_completion = openai.ChatCompletion.create(engine=chat_engine, messages=[{"role":"user", "content":prompt}])
    chat_completion = chat_complete([{"role": "user", "content": prompt}])
    return chat_completion


###################################################################################################################################################

def buyers_guide_content(username, jobtitle, compname, compto, deptbuget, country, noemp, interest, email):
    dic_criteria = {}
    dic_criteria[
        'Employee Experience & Engagement'] = 'capture, measure and analyse employee experience and engagement. Offering solutions to improve employee morale and satisfaction, enhance productivity, increase retention and improve communication.'
    dic_criteria[
        'Talent Acquisition'] = 'identifying organisational recruitment needs, source qualified candidates and hiring the best talent. It should streamline candidate search and recruitment processes and enhance candidate experience.'
    dic_criteria[
        'Learning & Development'] = 'create and deliver development courses, tracking learner progress, and automate learning. They should provide a centralized platform for leaders to track l&d progress, provide a range of different training materials and resources, provide powerful analytics and streamline learning processes to save time.'
    dic_criteria[
        'Transformation & change'] = 'Evolve business operations and digitally transform processes.  Provide the tools, guidance, and resources needed to effectively manage and implement change initiatives within their organization. They should also provide insights into organizational culture.'
    dic_criteria[
        'Culture & values'] = 'encourage collaboration, foster inclusivity, track and monitor employee feedback and help build a healthy, thriving workplace. They should identify, prioritize, and address key business challenges by providing a comprehensive view of my organization’s culture and values, allow leaders to access real-time, actionable insights into employee sentiment, engagement, and performance.'
    dic_criteria[
        'Diversity & Inclusion'] = 'ensure a diverse workforce through the creation of diverse talent pipelines, provide DE&I analytics and metrics, boost employee engagement and promote inclusion across the workforce.'
    dic_criteria[
        'Organisational Development & Effectiveness'] = 'improve the efficiency of project management, enabling open and effective communication, improve processes and systems related to performance management, talent management, diversity and employee wellness.'
    dic_criteria[
        'Rewards & Benefits'] = 'create, manage, and implement employee reward and benefit programs, motivate and engage employees, attract and retain top talent, and create a more positive, productive work environment.'
    dic_criteria[
        'Talent Management & Performance'] = 'provide valuable insights on employee performance, identify areas of improvement, reward outstanding performance, and provide employees with feedback to assist them in their development'
    dic_criteria[
        'Analytics & Data'] = 'providing actionable insights into employee performance, engagement, recruitment, and turnover. Help me better understand my workforce, identify trends and opportunities, and make informed decisions to improve the organisation’s overall performance.'

    job_title = jobtitle
    company_size = noemp
    company_name = compname
    interest_area = interest
    
    # Check if the user's interest area is in the dictionary keys
    if interest_area in dic_criteria:
        insert_criteria = dic_criteria[interest_area]
    else:
        # If the interest area is not in the dictionary, add it as a new key with a default prompt
        conversation = [
            {"role": "system", "content": "You are an AI-powered technology vendor comparison tool. Your purpose is to understand a user's key challenge area and generate a detailed report that provides an assessment of the best available HR technologies that could be purchased to help solve these challenges."},
            {"role": "user", "content": f"Provide a short description on {interest_area}."}
        ]
        generated_description = chat_complete(conversation)
        print("completed prompt defined interest area")
        dic_criteria[interest_area] = generated_description
        insert_criteria = generated_description
    
    budget = deptbuget
    country = country
    app_brand = "HR Director"
    app_brand_link = 'www.hrdconnect.com'

    # {

    # get all parameters
    # prompts
    conversation = [
        {"role": "system",
         "content": "You are an AI-powered technology vendor comparison tool. Your purpose is to understand a user's key challenge area and generate a detailed report that provides an assessment of the best available HR technologies that could be purchased to help solve these challenges. The report must also take in important demographic contextual information, as well as to assess the technologies against one another in order to help determine the best fit for each individual user. Your audience are senior business leaders and so the tone must be professional, grammatically structured for sophisticated readers and provide high-level information that experienced industry professionals would find useful. Do not use any keywords or identifiers which tells us this report is generated by AI. Use British English and business vocabulary."}
    ]

    prompt = {}
    prompt1 = f"I am a {job_title} at {company_name}, in {country}. The size of the company (number of employees) is {company_size}. I want to learn more about {interest_area}. Provide me with the 5 best and most relevant solutions that will help \"{insert_criteria}\". Responses must explicitly reference the name of the solution rather than a generic description. The solutions must be currently listed on their domestic corporate filing databases, including the SEC in the US or Companies House in the UK. If you are unable to find the solutions on the SEC or Companies House, regenerate a different solution."
    prompt2 = f"Provide a short description of these solutions."
    prompt3 = f"Regenerate, results should exclude solution providers considered as market leaders. Examples would include, but not limited to, Workday, ADP, SAP, Oracle, Cornerstone OnDemand, BambooHR and Paycor. Focus on identifying smaller/medium size providers who are disrupting the market with new technologies or solution."
    prompt4 = f"In bullet points, provide a summary of each of these against: cost (provide a range, my budget is {budget}), integration capabilities (against some examples of core HR systems) and scalability (number of employees specifically)"
    prompt5 = f"Regenerate, provide a specific estimated cost range based on what you know about me, my organisational needs as well as any pricing information you can provide from the vendors websites."
    prompt6 = f"Knowing what you know about my role and organization, provide me with a rating out of 10 for how suitable each of these solutions is. You must also include an explanation of around 40 words detailing why this product would be beneficial and an example what impact it might have to my business."
    prompt7 = f"Regenerate, answer must include both a score out of 10 and a description offering rationale for why it would be relevant."
    prompt8 = f"Provide me a peer review from G2 and Gartner insights, giving a score out of 5. It should include a quote from one of the reviewers. Do not reference either G2 or Gartner. If you can't do this, just give me a peer review that is credible. You must give me at least some review for these solutions."
    prompt9 = f"Provide a short summary on these solutions, comparing against one another. It should focus around these key areas: customisation, integration capabilities, content creation, reporting and analytics, scalability and flexibility, compliance and security, support and customer service and pricing. It is important to directly assess each platform against each criterion, detail strengths and weaknesses, and offer views as to which are better/more suitable for me based on the criteria and what you know about me. It should be at least 200 words, with a short introduction, a bullet for each criterion, and a summary in which the best platform is recommended."
    prompt10 = f"Write an introduction for a report for {app_brand}. The report should focused on {interest_area} and why it is an important area for a {app_brand} to consider for their business. The report should be informative and provide high-level strategical insights for a senior, experienced audience who have great depth of knowledge in the sector. It should include key industry trends, new technologies that are impacting the area as well as an overview of what {app_brand} should be considering for their businesses. Include reference to the specific challenges leaders face as reported in {app_brand_link} and the specific, explicit technologies that might help unlock these challenges. It should not feature any definitions of the subject matter, instead build upon existing knowledge of the challenge area."
    prompt11 = f"What to look for? What are meaningful metrics for {interest_area} considering I am a {job_title} at {company_name}, in {country}. The size of the company (number of employees) is {company_size}?"

    ###############################  run complete gpt code ########################################################################
    
    try:
        conversation.append({"role": "user", "content": prompt1})
        chat_completion = chat_complete(conversation)
        print(f"completed 1")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt2})
        chat_completion = chat_complete(conversation)
        print("completed 2")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt3})
        chat_completion = chat_complete(conversation)
        print("completed 3")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt4})
        chat_completion = chat_complete(conversation)
        print("completed 4")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt5})
        chat_completion = chat_complete(conversation)
        print("completed 5")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt6})
        chat_completion = chat_complete(conversation)
        print("completed 6")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt7})
        chat_completion = chat_complete(conversation)
        print("completed 7")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt8})
        chat_completion = chat_complete(conversation)
        print("completed 8")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt9})
        chat_completion = chat_complete(conversation)
        print("completed 9")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt10})
        chat_completion = chat_complete(conversation)
        print("completed 10")

        conversation.append({"role": "assistant", "content": chat_completion})
        conversation.append({"role": "user", "content": prompt11})
        # chat_completion = chat_complete(conversation)
        print("completed 11")

        conversation.append({"role": "assistant", "content": chat_completion})
    ########################################################################################################################

        # create a response dictionary
        response = {}
        # get intro
        introduction = generate_intro(conversation)

        # get solutions TODO remove this
        # solutions = generate_solutions(conversation)
        # get table content json
        table_data = generate_json_for_table(conversation)  # TODO might have to use eval function
        # get comparison
        comparison = generate_comparison(conversation)

        response['introduction'] = introduction
        # response['solutions_metrics'] = solutions  # TODO remove this
        response['table_content'] = table_data
        response['comparative_analysis'] = comparison

        return response

    except Exception as e:
        return {"Exception": str(e)+"\n\n Please try to resend the details again in sometime"}


################################################################################################


## unique id
def create_unique_id():
    """

    :param
    :return: Unique ID
    """
    # Get the current timestamp
    timestamp = int(time.time())

    # Define the characters to choose from
    characters = string.ascii_uppercase + string.digits

    # Generate a random 4-character code
    random_code = ''.join(random.choices(characters, k=4))

    unique_id = "-".join([str(timestamp), str(random_code)])

    return unique_id


## upload the json to s3 with unique id name. json

def save_Dict_to_json(dictionary, file_name):
    """

    :param dictionary: Dict,  dictionary that is to be saved
    :param file_name: str, file name
    :return: .json file saved in Json/ Folder
    """

    # Create the directory if it doesn't exist
    folder_path = "Json/"
    os.makedirs(folder_path, exist_ok=True)

    file_path = folder_path + file_name

    with open(file_path, "w") as json_file:
        json.dump(dictionary, json_file)



# function accepting all parameters

def set_user():
    i=0
    response = {}
    for member in sl:
        username = " ".join([member[2], member[1]])
        jobtitle = member[3]
        compname = member[4]
        compto = member[10]
        deptbuget = member[13]
        country = member[7]
        noemp = member[8]
        interest = member[11]
        email = member[5]
        user_info = {"firstName": member[2], "lastName": member[1], "title": jobtitle, "company": compname, "country": country, "No_of_Employees__c_contact": noemp, "Company_Turnover__c": compto, "learningAreas": interest, "departmentBudget": deptbuget, "email": email}
    # domain = data.get('domain')
        bg_content = buyers_guide_content(username, jobtitle, compname, compto, deptbuget, country, noemp, interest, email)

        if "Exception" in bg_content: print("OpenAI error occured ");continue

    # add user_info here instead giving it in buyers guide content function
        response.update({"User_submit_info": user_info})

    # add bg content
        response.update({"bg_content": bg_content})



    ## fucntion from Yash and santosh,
    # input: the response
    # output: saved docx file name
        docx_folder = "HRDData/"
        pdf_folder = "HRDPdf/"
    # Create the directories if it doesn't exist
        os.makedirs(docx_folder, exist_ok=True)
        os.makedirs(pdf_folder, exist_ok=True)

        docx_file_name = save_docx_from_response(response, docx_folder)

    ## convert docx to PDF
        convert_docx_to_pdf(docx_folder+docx_file_name , pdf_folder+docx_file_name.replace(".docx", ".pdf"))

    

################################################################################################################################

def main():
    set_user()
main()

