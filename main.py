import re
import PyPDF2
from openpyxl import load_workbook, Workbook
from tkinter import Tk, Label, Button, filedialog
from tkinter.ttk import Progressbar
import time
import os
import pdfplumber
from scripts.exc import excluded_phrases
from scripts.save_in_template import save_in_template

# Global variables
pdf_path = ''
pdf_name = ''
start_time = 0

# Function to search for the first date and project name in a PDF file and save them in Excel

def search_data():
    global pdf_path, pdf_name, start_time

    # Open the PDF file
    with pdfplumber.open(pdf_path) as pdf:
        # Create an Excel workbook and select the first sheet
        workbook = Workbook()
        sheet = workbook.active

        # Write headers in Excel
        sheet['A1'] = 'Meeting type'
        sheet['B1'] = 'Meeting Date'
        sheet['C1'] = 'Project Name'
        sheet['D1'] = 'Applicant'
        sheet['E1'] = 'Project Location'
        sheet['F1'] = 'Parcel'
        sheet['G1'] = 'Building Size'
        sheet['H1'] = 'Land Size'
        sheet['I1'] = 'Propose Project'
        sheet['J1'] = 'Existing Used'
        sheet['K1'] = 'Propose Zoning'
        sheet['L1'] = 'Current Application Status'
        sheet['M1'] = 'Comments'

        # Variables to store the results
        meeting_type = ''
        meeting_date = ''
        project_names = []
        parcel_numbers = []
        project_locations = []
        building_sizes = []
        land_sizes = []
        application_status = ''
        applicants = []
        proposals = []

        meeting_type = ''

        # After extracting the text from the PDF
        text = ""
        for page in pdf.pages:
            text += page.extract_text()

        # Check if the phrase "PLANNING COMMISSION" is present in any variant

        meeting_regex = r'PLANNING COMMISSION|Planning and Housing Commission'
        meeting_match = re.search(meeting_regex, text, re.I | re.M)
        if meeting_match: meeting_type = "Planning Commission Regular Meeting"
        else: meeting_type = "-"

        date_regex = r'\b\d{4}-\d{2}-\d{2}\b|\b\d{1,2}/\d{1,2}/\d{2,4}\b|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2}(?:st|nd|rd|th)?,? \d{2,4}\b'

        # Find all dates in the PDF

        all_dates = re.findall(date_regex, text)

        # Get the first date found in the document

        meeting_date = all_dates[0] if all_dates else '-'

        # Search for project names that match the regex based on the document content

        # Expresiones regulares existentes
        project_name_regexes = [
            r'PEN\d{2}-\d{4}',
            r'(?:PPM\d{4}-\d{4}|PM\s\d+|PC\s\d{2}-\d{4}|PP\d{4}-\d{4})|GPA\d{4}-\d{4}|CUP\d{4}-\d{4}',
            r'\d{4}-\d{2}',
            r'\d{1,2}-\d{3}',
            r'No\. \d{6} | \d{3}',
            r'No\. \d{2}-\d{3}(?:-\d{1,4})?',
            r'No\. \d{2}-\d{2}'
        ]

        # Variable para almacenar las coincidencias de los nombres de proyectos
        project_name_matches = []

        # Bucle para buscar coincidencias con las expresiones regulares
        for project_name_regex in project_name_regexes:
            matches = re.findall(project_name_regex, text, re.I)
            project_name_matches.extend(matches)

        # Eliminar duplicados y asignar valor '-' si no hay coincidencias
        project_names = list(set(project_name_matches)) if project_name_matches else ['-']

        # Search for parcel information in the format "{###-###-###}"

        parcel_number_regex = r"\d\d\d-\d\d\d-\s?\d\d\d"
        parcel_number_matches = re.findall(parcel_number_regex, text)

        if not parcel_number_matches:
            pattern = r"Tract(?:\s+Map)?\s+(\d+)"
            match = re.search(pattern, text, re.I)
            if match:
                tract_number = match.group(1)
                parcel_numbers = ['Tract No. ' + tract_number]
            else:
                # Use \d{4}-\d{3}-\d{3} pattern
                parcel_number_regex = r"APN:([\s\S]*?)\d{4}-\d{3}-?\s?\d{3}"
                parcel_number_matches = re.findall(parcel_number_regex, text, re.M)
                if parcel_number_matches:
                    parcel_numbers = list(set(parcel_number_matches))
                else:
                    parcel_numbers = []
        else:
            parcel_numbers = list(set(parcel_number_matches))

        # Search for all locations that match coordinates or any name in any format representing a physical place

        location_regex = r'Location:\s*(.*?)\s*[\n\r]'
        location_matches = re.findall(location_regex, text, re.I | re.M)
        site_regex = r'Project Site:\s*(.*?)\s*[\n\r]'
        site_matches = re.findall(site_regex, text, re.I | re.M)
        located_regex = r'located\s*at\s*([\w\d\s.-]+)'
        located_matches = re.findall(located_regex, text, re.I | re.M)
        located_on_regex = r'located on\s*([\w\d\s.-]+)'
        located_on_matches = re.findall(located_on_regex, text, re.I | re.M)
        located_in_regex = r'located in\s*([\w\d\s.-]+)'
        located_in_matches = re.findall(located_in_regex, text, re.I | re.M)
        locations = []
        for location in location_matches:
            if re.search(r'(\d+(\.\d+)?),\s*(\d+(\.\d+)?)', location):
                # If the location matches coordinates, add it to the locations list
                locations.append(location)
            else:
                # If the location doesn't match coordinates, check if it contains any name representing a physical place
                if re.search(r'\b[A-Za-z\s]+\b', location):
                    locations.append(location)

        for location in site_matches:
            if re.search(r'(\d+(\.\d+)?),\s*(\d+(\.\d+)?)', location):
                # If the location matches coordinates, add it to the locations list
                locations.append(location)
            else:
                # If the location doesn't match coordinates, check if it contains any name representing a physical place
                if re.search(r'\b[A-Za-z\s]+\b', location):
                    locations.append(location)

        for location in located_matches:
            if re.search(r'(\d+(\.\d+)?),\s*(\d+(\.\d+)?)', location):
                # If the location matches coordinates, add it to the locations list
                locations.append(location)
            else:
                # If the location doesn't match coordinates, check if it contains any name representing a physical place
                if re.search(r'\b[A-Za-z\s]+\b', location):
                    locations.append(location)

        for location in located_on_matches:
            if re.search(r'(\d+(\.\d+)?),\s*(\d+(\.\d+)?)', location):
                locations.append(location)
            else:
                if re.search(r'\b[A-Za-z\s]+\b', location):
                    locations.append(location)
        
        for location in located_in_matches:
                    if re.search(r'(\d+(\.\d+)?),\s*(\d+(\.\d+)?)', location):
                        locations.append(location)
                    else:
                        if re.search(r'\b[A-Za-z\s]+\b', location):
                            locations.append(location)
                    
        project_locations = list(set(locations)) if locations else ['-']

        # Search for the application status ("Approved" or "Approval") in any format or variation

        application_status_regex = r'\b(APPROVED|APPROVAL|APPROVE)\b'
        application_status_match = re.search(application_status_regex, text, re.I)

        if application_status_match:
            application_status = "APPROVED"
        else:
            application_status = '-'

       # Search for applicant names
        applicant_regex = r"Applicant: ([A-Z][a-z]+ [A-Z][a-z]+)"
        applicant_matches = re.findall(applicant_regex, text, re.I | re.M)
        if not applicant_matches:
            applicant_regex = r"\b[A-Z][a-z]+ [A-Z][a-z]+\b(?:\sDevelopment Group)?"
            applicant_matches = re.findall(applicant_regex, text)
        if not applicant_matches:
            applicant_regex = r"Owner:\s*([^.,\n]+)"
            applicant_matches = re.findall(applicant_regex, text)
        unique_applicants = list(set(applicant_matches))
        filtered_applicants = [applicant for applicant in unique_applicants if applicant not in excluded_phrases]
        applicants = filtered_applicants or ['-']

        # Search for building size with the format if not found in the previous format

        building_size_regex = r'\(?(\d{1,3}(?:,\d{3})*(?:\.\d+)?)\)?\s*(SQUARES?\s*FEETS?|sq.\s*ft|SF|SQUARES?\s*FooTS?)'
        building_size_match = re.findall(building_size_regex, text, re.I)
        building_sizes = list(set([size[0].strip('()') + ' SF' for size in building_size_match])) if building_size_match else ['-']

        # Search for land sizes in acres with the formats "2.6 acre", "33.57-acre", or "18.49- acre site"

        land_size_regex = r'(\d+(?:\.\d+)?)\s*(?:-)?\s*acre\s*(?:site)?'
        land_size_matches = re.findall(land_size_regex, text, re.I)
        land_sizes = list(set([size + ' acre' for size in land_size_matches])) if land_size_matches else ['-']

        # Search for proposals after "Proposal" and save the text until the first period

        moreno_proposal_regex = r'(?:Proposal|Proposed Project|Proposal: |Proposed Project: )\b([^.:]*\d+(?:\.\d+)?[^.]*)\.'
        corona_proposal_regex = r'PUBLIC HEARING\s*-\s*([^*]+?)\bApplicant:'
        elsinore_proposal_regex = r'\bID#\s\d{2}-\d{3}\b\s*((?:(?!(?:Attachments\b|.*\bcoronavirus\b)).)*)'
        hemet_proposal_regex = r'PROJECT SUMMARY:\s*(.*?\.)'
        stfe_proposal_regex = r'a request\b([^.]*)\.'
        malibu_proposal_regex = r'Recommended Action\b([^.]*)\.'
        gabriel_proposal_regex = r'The proposed project\s.([^\.]+\.[^\.]+\.[^\.]+\.[^\.]+(?:\.[^\.]+)?)'
        puente_proposal_regex = r'CONSIDERATION\s+(.*?)(?=\n\n|[A-Z]{2,})'

        proposals = re.findall(moreno_proposal_regex, text, re.S)
        if not proposals:
            proposals = re.findall(corona_proposal_regex, text, re.S)
        if not proposals:
            proposals = [match.group(1).strip() for match in re.finditer(elsinore_proposal_regex, text, re.S | re.IGNORECASE) if 'coronavirus' not in match.group(1).lower()]
        if not proposals:
            hemet_matches = re.findall(hemet_proposal_regex, text, re.S | re.M)
            proposals = [match.strip() for match in hemet_matches]
        if not proposals:
            stfe_matches = re.findall(stfe_proposal_regex, text, re.S | re.I | re.M | re.U)
            proposals = [match.strip() for match in stfe_matches]
        if not proposals:
            stfe_matches = re.findall(malibu_proposal_regex, text, re.S | re.I | re.M)
            proposals = [match.strip() for match in stfe_matches]
        if 'San Gabriel' in text:
            gabriel_matches = re.findall(gabriel_proposal_regex, text, re.I | re.M)
            proposals = [match.strip() for match in gabriel_matches]
        if 'La Puente' in text:
            puente_matches = re.findall(puente_proposal_regex, text, re.M)
            proposals = [match.strip() for match in puente_matches]

        captura = proposals[0] if proposals else '-'

        # Search for existing use

        existing_used = []

        # Search for the word "existing" followed by the complete phrase until the first period

        existing_regex = r'(?i)\bexisting\b[^.]*'
        existing_matches = re.findall(existing_regex, text)
        existing_used = existing_matches if existing_matches else ['-']

        # Search for propose zoning

        propose_zoning_regex = r'(?i)\b(?:proposal|PUBLIC HEARING)\b.*?\b(Construction|subdivide|Expansion|Merge|Remodel|Subdivition(?:\s+and)?(?:\s+Construction)?|Development|Demolition(?:\s+and\s+construction)?)\b'
        propose_zoning_matches = re.findall(propose_zoning_regex, text, re.I | re.S | re.M)
        propose_zoning = propose_zoning_matches if propose_zoning_matches else '-'

        # Write the results to Excel

        sheet['A2'] = meeting_type
        sheet['B2'] = meeting_date
        sheet['C2'] = '; '.join(project_names)
        sheet['D2'] = ', '.join(applicants)
        sheet['E2'] = '; '.join(project_locations)
        sheet['F2'] = '; '.join(parcel_numbers)
        sheet['G2'] = '; '.join(building_sizes)
        sheet['H2'] = '; '.join(land_sizes)
        sheet['I2'] = '; '.join(proposals)
        sheet['J2'] = '; '.join(existing_used)
        sheet['K2'] = '; '.join(propose_zoning)
        sheet['L2'] = application_status

        # Save the Excel file

        excel_path = f'{pdf_name}_results.xlsx'
        workbook.save(excel_path)

        # Save the data in the first empty row of the "COPIA PLANTILLA" file

        # Create a list with the collected data

        data = [
            meeting_type,
            meeting_date,
            '; '.join(project_names),
            '; '.join(applicants),
            '; '.join(project_locations),
            '; '.join(parcel_numbers),
            '; '.join(building_sizes),
            '; '.join(land_sizes),
            '; '.join(proposals),
            '; '.join(existing_used),
            '; '.join(propose_zoning),
            application_status,
        ]

        #Call the function to save the data in the excel template
        
        save_in_template(data, 'COPIA PLANTILLA.xlsx')

        # Call the function to save the data in MongoDB

        # Display completion message and total execution time
        lbl_messagge.config(
            text=f'The search has been finished. Total time: {time.time() - start_time:.2f} segundos')

        # Open the Excel file after saving
        os.startfile(excel_path)

       # Display results in the console
        print(f"Meeting Type: {meeting_type}")
        print(f"Meeting Date: {meeting_date}")
        print(f"Project Names: {project_names}")
        print(f"Applicants: {applicants}")
        print(f"Project Locations: {project_locations}")
        print(f"Parcel Numbers: {parcel_numbers}")
        print(f"Building Sizes: {building_sizes}")
        print(f"Land Sizes: {land_sizes}")
        print(f"Proposals: {proposals}")
        print(f"Application Status: {application_status}")
        print(f"Existing / Used: {existing_used}")
        print(f"Propose Zoning: {propose_zoning}")
        print("Results saved to Excel file.")

    # Enable the search button after completing the task
    btn_search.config(state='normal')

# Function to select a PDF file

def select_pdf():
    global pdf_path, pdf_name, start_time

    # Open the file selection dialog
    pdf_path = filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])

    # Get the name of the selected PDF file
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

    # Disable the search button while the task is running
    btn_search.config(state='disabled')

    # Display start search message
    lbl_messagge.config(text='Searching data in the file...')

    # Save the start time of execution
    start_time = time.time()

    # Call the function to search for data in the PDF file
    search_data()


# GUI configuration
root = Tk()
root.title('Metro Data Tool')
root.geometry('300x200')

# Estilo para los labels
label_style = {'font': ('Arial', 12)}

# Estilo para el botón
button_style = {'font': ('Arial', 12), 'bg': '#283593', 'fg': 'white', 'activebackground': '#1a237e', 'activeforeground': 'white'}

# Estilo para la información del programa
info_style = {'font': ('Arial', 10, 'italic'), 'fg': 'gray'}

lbl_instrucciones = Label(root, text='Select a PDF file:', **label_style)
lbl_instrucciones.pack(pady=30)

btn_search = Button(root, text='Search', command=select_pdf, **button_style)
btn_search.pack()

lbl_message = Label(root, text='', **label_style)
lbl_message.pack(pady=10)

lbl_info = Label(root, text='Version 1.0.1', **info_style)
lbl_info.pack(side='bottom')

root.mainloop()