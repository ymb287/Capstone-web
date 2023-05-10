import streamlit as st
import pandas as pd
import xlwings as xw
import base64
import tempfile
import os

# Set the page up
st.set_page_config(layout="wide")
st.markdown("""<style>#MainMenu {visibility: hidden;}footer {visibility: hidden;}</style>""", unsafe_allow_html=True) 

def image_to_base64(img_path):
    with open(img_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode('utf-8')


st.markdown(
    """
    <style>
        .top-right-image {
            position: absolute;
            top: -5rem;
            right: 0rem;
            max-height: 50px;
            max-width: 50px;
        }
        .shifted-up {
            margin-top: -5rem;
        }
    </style>
    
    """,
    unsafe_allow_html=True,
)


image_base64 = image_to_base64("hsg.jpg")

def excel_calculation(input1,input2,input3,input4,input5,input6,input7,input8,input9,input10):
    # Load the Excel workbook
    workbook = xw.Book('tool_detective_business_case.xlsx')


    # Select the input sheet and output sheet by name
    input_sheet = workbook.sheets['Inputfaktoren']
    output_sheet = workbook.sheets['Dashboard']

    # Update the value of the input cell
    input_sheet.range('C6').value = input1
    input_sheet.range('C7').value = input2
    input_sheet.range('C10').value = input3
    input_sheet.range('C11').value = input4
    input_sheet.range('C12').value = input5
    input_sheet.range('C15').value = input6
    input_sheet.range('C16').value = input7
    input_sheet.range('C17').value = input8
    input_sheet.range('C20').value = input9
    input_sheet.range('C21').value = input10



    # Recalculate the formulas in the workbook
    workbook.app.calculate()

    # Get the updated output values from multiple cells in the output sheet
    output_data1 = output_sheet.range('C85:Z85').value
    output_data1 = [round(num, 0) for num in output_data1]
    output_data2 = output_sheet.range('C89:Z89').value
    output_data2 = [round(num, 0) for num in output_data2]



    # Save the updated workbook to a temporary file
    temp_dir = tempfile.mkdtemp()
    temp_file = os.path.join(temp_dir, 'updated_tool_detective_business_case.xlsx')
    workbook.save(temp_file)

    # Close the workbook
    workbook.close()

    # Save the temporary file path in the session state
    st.session_state.temp_file = temp_file

    # Check if dataframe is not empty before rounding values
    st.session_state.output_data1 = output_data1
    st.session_state.output_data2 = output_data2

    return output_data1,output_data2



def filedownload(temp_file):
    with open(temp_file, "rb") as f:
        base64_data = base64.b64encode(f.read()).decode("utf-8")
        href = f'<a href="data:application/octet-stream;base64,{base64_data}" download="updated_tool_detective_business_case.xlsx">Download Updated Excel File</a>'
    return href

with st.sidebar: 
    st.image("https://www.onepointltd.com/wp-content/uploads/2020/03/inno2.png")
    st.title("Business Tool")
    choice = st.radio("Navigation", ["Parmeters","Output","Download"])
    st.info("This project application helps you developing XYZ.")

if choice == "Parmeters":
    st.markdown("<h1 class='shifted-up'>Parameter Input</h1>", unsafe_allow_html=True)
    st.markdown(f'<img src="data:image/jpeg;base64,{image_base64}" class="top-right-image" />', unsafe_allow_html=True,)


    with st.form("Form1"):

        col1 = st.columns(1)
        with col1[0]:
            st.write("<div style='text-align: center; font-weight: bold;'>Stückzahl</div>", unsafe_allow_html=True)


        col2, col3 = st.columns([1,1])
        with col2:
            st.write("Anzahl der Zerspanungsapparate")
            input1 = st.number_input("", value=33, key='input1')
        with col3:
            st.write("Anzahl Zerspanungswerkzeuge welche Abnutzungskontrollen unterstehen")
            input2 = st.number_input("", value=387, key='input2')
        st.write("<hr>", unsafe_allow_html=True)


        col4 = st.columns(1)
        with col4[0]:
            st.write("<div style='text-align: center; font-weight: bold;'>Monetäre Beträge</div>", unsafe_allow_html=True)

        col5, col6 , col7 = st.columns([1,1,1])
        with col5:
            st.write("Stundenlohn pro Mitarbeiter")
            input3 = st.number_input("", value=1000, key='input3')
        with col6:
            st.write("Kosten von Ausschussteilen durch abgenutzte Zerspanungswerkzeuge pro Monat")
            input4 = st.number_input("", value=1000, key='input4')
        with col7:
            st.write("Materialaufwand für Zerspanungswerkzeuge pro Jahr")
            input5 = st.number_input("", value=387, key='input5')
        st.write("<hr>", unsafe_allow_html=True)


        col8 = st.columns(1)
        with col8[0]:
            st.write("<div style='text-align: center; font-weight: bold;'>Zeitangaben in Stunden und Zerspanungswerkzeug pro Monat</div>", unsafe_allow_html=True)
        
        col9, col10 , col11 = st.columns([1,1,1])
        with col9:
            st.write("Manuelle Werkzeugkontrolle")
            input6 = st.number_input("", value=0.1, key='input6')
        with col10:
            st.write("Für Nacharbeit von Produktionserzeugnissen, durch abgenutzten Zerspanungswerkzeugen")
            input7 = st.number_input("", value=4, key='input7')
        with col11:
            st.write("Für die Werkzeugqualifizierung von Zerspanungswerkzeugen ")
            input8 = st.number_input("", value=387, key='input8')
        st.write("<hr>", unsafe_allow_html=True)


        col12 = st.columns(1)
        with col12[0]:
            st.write("<div style='text-align: center; font-weight: bold;'>Prozentangaben </div>", unsafe_allow_html=True)

        col13, col14 = st.columns([1,1])
        with col13:
            st.write("Durchschnittlicher Verschleiss eines Zerspanungswerkzeugs vor Entsorgung in %")
            input9 = st.number_input("", value=60, key='input9')
            input9= input9/100
        with col14:
            st.write("Durchschnittliche Preissteigerung für Material in Zerspanungswerkzeugen (2022 - 2024E) in %")
            input10 = st.number_input("", value=387, key='input10')
            input10= input10/100

        calculate_button = st.form_submit_button("Calculate")

        if calculate_button:
            with st.spinner("Calculating..."):
                pass
            output_data1, output_data2 = excel_calculation(input1,input2,input3,input4,input5,input6,input7,input8,input9,input10)


if choice == "Output":
    st.markdown("<h1 class='shifted-up'>Output</h1>", unsafe_allow_html=True)
    st.markdown(f'<img src="data:image/jpeg;base64,{image_base64}" class="top-right-image" />', unsafe_allow_html=True,)

    if 'output_data1' not in st.session_state:
        st.write("Senario1")
        st.write("No output data available yet.")
    if 'output_data2' not in st.session_state:
        st.write("Senario2")
        st.write("No output data available yet.")

    else:
        columns = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]
        # Define the data as two lists
        data = [st.session_state.output_data1, st.session_state.output_data2]
        # Create a DataFrame from the data and columns
        df = pd.DataFrame(data, index=['Scenario 1', 'Senario 2'], columns=columns)
        st.write(df)




if choice == "Download":
    st.markdown("<h1 class='shifted-up'>Download</h1>", unsafe_allow_html=True)
    st.markdown(f'<img src="data:image/jpeg;base64,{image_base64}" class="top-right-image" />', unsafe_allow_html=True,)
    if 'temp_file' in st.session_state:
        st.markdown(filedownload(st.session_state.temp_file), unsafe_allow_html=True)
    else:
        st.write("No updated Excel file available yet.")
