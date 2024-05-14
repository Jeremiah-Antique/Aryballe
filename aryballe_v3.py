import streamlit as st, pandas as pd, plotly.express as px, io
from pandas import ExcelWriter
from pandas import ExcelFile
from streamlit.logger import get_logger
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
import plotly.io as pio
import plotly.offline as pyo
from PIL import Image
import plotly.graph_objs as go
import webbrowser
import datetime
from xlsxwriter import Workbook
import xlsxwriter

#from aryballe_HT_Data import data_uploader
LOGGER = get_logger(__name__)
title_container = st.container(border=False)
header_container = st.container(border=True)
report_container = st.container(border=True)

def convert_unix_to_time_duration(df, timestamp):
    # Ensure all entries are numeric, converting from string if necessary
    df[timestamp] = pd.to_numeric(df[timestamp], errors='coerce')

    # Drop any rows where the timestamp could not be converted to a number
    df.dropna(subset=[timestamp], inplace=True)

    # Reset the DataFrame's index to ensure consistency after dropping rows
    df.reset_index(drop=True, inplace=True)

    # Check if DataFrame is empty after dropping rows
    if df.empty:
        print("All rows have invalid timestamps and have been dropped.")
        return pd.Series([]) #return a series to avoid assignment errors

    # Convert the first valid timestamp from milliseconds to seconds
    start_time = datetime.datetime.fromtimestamp(df.iloc[0][timestamp] / 1000)

    # Calculate the time duration from the start time for each timestamp
    return df[timestamp].apply(
        lambda x: (datetime.datetime.fromtimestamp(x / 1000) - start_time).total_seconds() / 60
    )

def create_bar(df,x,y,title,sort_column,avg):
    fig = px.bar(df, x=x, y=y, title=title,
                                facet_col='run_id', barmode='group',
                                category_orders={"cycle": sorted(df[sort_column].unique())})
    if avg is not None:
        fig.add_hline(y=avg,line_dash="dash", line_color="red")
    st.plotly_chart(fig)
    return fig

def aryballe_average(df,cycle_col, filename):
    cycle_options = []
    df[cycle_col] = pd.to_numeric(df[col_type],errors='coerce')
    cycle_options.extend(df[cycle_col].unique().round()) 
    cycle_options = sorted(set(cycle_options), key=lambda x: int(x)) #removing duplilcates and sorting
    
    if cycle_options:   
        initial_cycle = st.selectbox('Select Initial Cycle',cycle_options,index=0,key=f"{filename}_initial")
        end_cycle = st.selectbox('Select End Cycle',cycle_options,index=len(cycle_options)-1,key=f"{filename}_end")
        st.write(f" You selected and initial cycle of {initial_cycle} and a final cycle of {end_cycle}")
        all_intensities = []
    
        # Filter DataFrame for selected cycles
        filtered_df = df[df[cycle_col].between(initial_cycle, end_cycle)]
        average_intensity = filtered_df['mean intensity'].mean().round(3)
        all_intensities.append(average_intensity)
    
    if all_intensities:
        overall_average = sum(all_intensities) / len(all_intensities)
        st.markdown(f"<p style='color: red; font-size: 24px;'>Average intensity from {initial_cycle} to {end_cycle} is {overall_average}</p>",unsafe_allow_html=True)
        avg_fig = create_bar(filtered_df,cycle_col,'mean intensity',f"Average intensity between {initial_cycle} and {end_cycle} for {filename}",cycle_col,overall_average)
        avg = filtered_df
    return avg,avg_fig,overall_average

def avg_intensity_time_range(df,col_type,filename):
        st.header(f"Average Intensity for {filename}",divider="red")
        if col_type == 'cycle':
            
            avg,avg_fig,overall_average = aryballe_average(df, 'cycle', filename)
            return avg,avg_fig
        elif col_type == 'Time':
          
            avg,avg_fig,overall_average = aryballe_average(df,'Time', filename)
            return avg,avg_fig
        else:
            st.write("No cycle data available for the selected cycles.")
            avg=[]
            avg_fig=[] 
            return avg,avg_fig

def broken_create_excel_report_with_plotly_fig_and_chart(fig, avg_fig, df, intensity_avg, filename='Aryballe report.xlsx'):
    # Initialize the byte variables to None to ensure they are defined in all execution paths
    img_bytes = None
    img_bytes_2 = None

    # Convert Plotly figure to a PNG image bytes if the figure is not None
    if fig is not None:
        img_bytes = fig.to_image(format="png")
        # Save the PNG file and store the path in session state
        fig_filename = f"{filename[:-5]}_fig.png"
        with open(fig_filename, "wb") as f:
            f.write(img_bytes)
        if 'image_paths' not in st.session_state:
            st.session_state['image_paths'] = {}
        st.session_state['image_paths']['fig_path'] = fig_filename

    # Convert the second Plotly figure to a PNG image bytes if the figure is not None
    if avg_fig is not None:
        img_bytes_2 = avg_fig.to_image(format="png")
        # Save the PNG file and store the path in session state
        avg_fig_filename = f"{filename[:-5]}_avg_fig.png"
        with open(avg_fig_filename, "wb") as f:
            f.write(img_bytes_2)
        if 'image_paths' not in st.session_state:
            st.session_state['image_paths'] = {}
        st.session_state['image_paths']['avg_fig_path'] = avg_fig_filename

        workbook = xlsxwriter.Workbook('chart.xlsx')
        worksheet = workbook.add_worksheet('Data')

        #add bold format
        bold = workbook.add_format({'bold': True})

        #write some headers
        worksheet.write('A1', 'Cycle/Minutes', bold)
        worksheet.write('B1', 'Intensity', bold)

        #start from first cell in sheet, row is 1 to be below header and column are zero index
        row = 1
        col = 0 
        #iterate over the date and write row by row
        for row_data in df.iterrows():
            worksheet.write(row, col, row_data['Minutes'])
            worksheet.write(row, col + 1, row_data['Intensity'])
            row+= 1

        #define last row
        last_row = len(df)+1 #+1 because of header is in row 1
        #define formula for categories (x) and values (y)
        categories_formula = ['Data', 0,0,4,0]
        #f'Data!$A$2:$A${last_row}'
        values_formula = ['Data', 0,1,4,0]
        #f'Data!$B$2:$B${last_row}'

        chart = workbook.add_chart({'type':'column'})
        #chart.add_series({'categories': categories_formula,'values': values_formula})

        # Set chart properties
        #chart.set_title({'name': 'Intensity by Cycle'})
        #chart.set_x_axis({'name': 'Cycle/Minutes'})
        #chart.set_y_axis({'name': 'Intensity'})

    # Insert the chart into the worksheet at position D1
        #worksheet.insert_chart('D1', chart)

        workbook.close()
    # The file is saved and closed by the context manager
    return filename

def create_excel_report_with_plotly_fig_and_chart(fig, avg_fig, df, intensity_avg, filename='Aryballe report.xlsx'):
    # Initialize the byte variables to None to ensure they are defined in all execution paths
    img_bytes = None
    img_bytes_2 = None

    # Convert Plotly figure to a PNG image bytes if the figure is not None
    if fig is not None:
        img_bytes = fig.to_image(format="png")
        # Save the PNG file and store the path in session state
        fig_filename = f"{filename[:-5]}_fig.png"
        with open(fig_filename, "wb") as f:
            f.write(img_bytes)
        if 'image_paths' not in st.session_state:
            st.session_state['image_paths'] = {}
        st.session_state['image_paths']['fig_path'] = fig_filename

    # Convert the second Plotly figure to a PNG image bytes if the figure is not None
    if avg_fig is not None:
        img_bytes_2 = avg_fig.to_image(format="png")
        # Save the PNG file and store the path in session state
        avg_fig_filename = f"{filename[:-5]}_avg_fig.png"
        with open(avg_fig_filename, "wb") as f:
            f.write(img_bytes_2)
        if 'image_paths' not in st.session_state:
            st.session_state['image_paths'] = {}
        st.session_state['image_paths']['avg_fig_path'] = avg_fig_filename

    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Convert DataFrame to Excel
        df.to_excel(writer, sheet_name='Data', index=False)
        intensity_avg.to_excel(writer, sheet_name='Intensity Average', index=False)

        # If image bytes are available, insert the image into the 'Data' worksheet
        if img_bytes is not None:
            workbook = writer.book
            worksheet = writer.sheets['Data']
            image_stream = BytesIO(img_bytes)
            worksheet.insert_image('G2', 'plotly_figure.png', {'image_data': image_stream})

        # If the second image bytes are available, insert the image into the 'Intensity Average' worksheet
        if img_bytes_2 is not None:
            workbook = writer.book
            worksheet = writer.sheets['Intensity Average']
            image_stream_2 = BytesIO(img_bytes_2)
            worksheet.insert_image('G2', 'plotly_figure_2.png', {'image_data': image_stream_2})

    # The file is saved and closed by the context manager
    return filename

def combine_excel_sheets(selected_reports,output_filename='Combined_Report.xlsx'):
    #create new excel writer object
    with pd.ExcelWriter(output_filename,engine='openpyxl') as writer:
        for report in selected_reports:
            #load each report
            xls = pd.ExcelFile(st.session_state.reports[report])
            #Write each sheet in the current report to the new file
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls,sheet_name=sheet_name)
                #ensure unique sheet names by appending report name and sheet
                unique_sheet_name = f"{report.split('.')[0]}_{sheet_name}"[:31] #excel sheet names limited to 31 char
                df.to_excel(writer,sheet_name=unique_sheet_name, index=False)
        
    return output_filename

def create_excel_bar_chart(df, x, y, title,sheet, sort_column, avg, filename):
    # Create a writer object using xlsxwriter and the specified file name
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Write the DataFrame to an Excel sheet
        df.to_excel(writer, sheet_name=sheet, index=False)



def sensory_panel(uploaded_file):
    hyphen = '-'
   
    def hyphen_string(value):
        #check for hyphen in string, split to keep only first string portion
        if hyphen in value:
            return (value.split(hyphen)[0])
        else:
            return ''
    def split_str(df_name):
        #split string applies Hyphen string to df, to extract value and then convert to numeric integer
        new_df = df_name.copy()
        new_df['HTS_split'] = df_name['HTS'].apply(hyphen_string)
        new_df['HTS'] = pd.to_numeric(new_df['HTS_split'], errors='coerce')
        
        # Apply hyphen_string function to 'Agree or Disagree' column
        try:
            new_df['agree_split'] = df_name['Agree or Disagree'].apply(hyphen_string)
            new_df['Agree or Disagree'] = pd.to_numeric(new_df['agree_split'], errors='coerce')
            # Drop temporary columns
            new_df.drop(columns=['agree_split'])
        except:
            print("Cannot apply agree or disagree split")
        try:
            new_df['lik_split'] = df_name['Like/Dislike?'].apply(hyphen_string)
            new_df['Like/Dislike?'] = pd.to_numeric(new_df['lik_split'], errors='coerce')
            # Drop temporary columns
            new_df.drop(columns=['lik_split'])
        except:
            print("Cannot apply Like/dislike split")
        return new_df.drop(columns=['HTS_split'])
    
    def Average_HTS(df_name,name):
        #drop empty rows
        new_df = df_name.dropna(axis = 0, how='any')
        new_df = split_str(new_df)
        #filter new_df by user input "name"
        filtered_df = []
        filtered_df = new_df[new_df['Fragrance'] == name]
        #Display filtered df on app page
        st.write(filtered_df)
        #average hot throw of fragrance sorted by "name" and round 2 decimals
        Avg_HTS = new_df[new_df['Fragrance'].str.strip() == name]['HTS'].mean().round(2)
        #generate bar graph by name and hot throw score
        fig= px.bar(filtered_df, x = 'Name', y= 'HTS',title= 'HT Scores by Fragrance')
        st.plotly_chart(fig)
        #return variables for report generation
        return fig,new_df,Avg_HTS,filtered_df
    
    def NHT_Average_HTS(df_name):

            # Group the data by 'Name' and 'Fragrance'.
        grouped_df = df_name.groupby(['Fragrance'])
        
        # Calculate the average 'HTS' for each grouping of 'Door' and 'Fragrance'.
        average_hts = grouped_df['HTS'].mean().round(2)

        # Create a new DataFrame with the results.
        data = pd.DataFrame({'Fragrance': average_hts.index.get_level_values(0),
                            'Average_HTS': average_hts})
        def reformat(df_name):
            #rename columns, then drop them and reset index to clean up df
            df_name =df_name[df_name['Fragrance'] !='Fragrance']
            df_name = df_name.rename(columns={'Fragrance': 'Frag'})
            df_name = df_name.drop(['Frag'], axis=1)
            df_name = df_name.reset_index()
            return df_name

        data = reformat(data)
        return data
    
    def run_NHT_average(df_name):
        new_df = df_name.dropna(axis = 0, how='any')
        st.write(new_df)
        new_df = split_str(new_df)
        HT_average = NHT_Average_HTS(new_df)
        st.write("Running Fragrance HT Average")
    
        st. write(HT_average)
        fig= px.bar(HT_average, x = 'Fragrance', y= 'Average_HTS',title= 'HT Scores by Fragrance')
        st.plotly_chart(fig)
        

        return fig,new_df,HT_average
    
    def Ap_avg(df_name,name):
        new_df = df_name.dropna(axis = 0, how='any')
        new_df = split_str(new_df)
        #new_df = new_df.drop(columns=['lik_split'])
        filtered_df = []
        filtered_df = new_df[new_df['Fragrance'] == name]
        st.write(filtered_df)
        Avg_ap = new_df[new_df['Fragrance'].str.strip() == name]['Agree or Disagree'].mean().round(2)
        fig= px.bar(filtered_df, x = 'Name', y= 'Agree or Disagree',title= 'Accuracy Scores by Name')
        st.plotly_chart(fig)
        return fig, new_df,Avg_ap,filtered_df
    
    def agree_avg(df_name):
        df_name = df_name.groupby(['Fragrance'])
        # Calculate the average 'HTS' for each grouping of 'Door' and 'Fragrance'.
        average_like = df_name['Agree or Disagree'].mean().round(2)
    
        # Create a new DataFrame with the results.
        data = pd.DataFrame({'Agree or Disagree': average_like.index.get_level_values(0),
                                'Average Agree or Disagree': average_like})
        def reformat(df_name):
            #rename columns, then drop them and reset index to clean up df
            df_name =df_name[df_name['Agree or Disagree'] !='Agree or Disagree']
            df_name = df_name.rename(columns={'Agree or Disagree': 'Like'})
            df_name = df_name.drop(['Like'], axis=1)
            
            df_name = df_name.reset_index()
            return df_name

        data = reformat(data)
        return data
    
    def run_agree_avg(df_name):
        new_df = df_name.dropna(axis = 0, how='any')
        st.write(new_df)
        new_df = split_str(new_df)
        agree_average = agree_avg(new_df)
        st.write("Running Fragrance Like/Dislike Average")
    
        st. write(agree_average)
        fig= px.bar(agree_average, x = 'Fragrance', y= 'Average Agree or Disagree',title= 'Average Accuracy by Fragrance')
        st.plotly_chart(fig)
        fig=go.Figure(fig)
        return fig,new_df,agree_average
    
    def run_frag_door_average(df_name):
        new_df = df_name.dropna(axis = 0, how='any')
        st.write(new_df)
        new_df = split_str(new_df)
        HT_average = D_Average_HTS(new_df)
        HT_average = HT_average[HT_average['Fragrance'] != "Fragrance"]
        st.write("Running Door HT Average")
        st. write(HT_average)
        fig= px.bar(HT_average, x = 'Fragrance', y= 'Average_HTS',color='Door',title= 'HT Scores by Fragrance',barmode='group')
        st.plotly_chart(fig)
        return fig,new_df,HT_average
        
    def D_Average_HTS(df_name):
            # Group the data by 'Name' and 'Fragrance'.
        grouped_df = df_name.groupby(['Door', 'Fragrance'])

        # Calculate the average 'HTS' for each grouping of 'Door' and 'Fragrance'.
        average_hts = grouped_df['HTS'].mean().round(2)

        # Create a new DataFrame with the results.
        data = pd.DataFrame({'Door': average_hts.index.get_level_values(0),
                            'Fragrance': average_hts.index.get_level_values(1),
                            'Average_HTS': average_hts})
        def reformat(df_name):
            #rename columns, then drop them and reset index to clean up df
            df_name = df_name.rename(columns={'Fragrance': 'Frag', 'Door': 'Door'})
            df_name = df_name.drop(['Door', 'Frag'], axis=1)
            df_name = df_name.reset_index()
            return df_name

        data = reformat(data)
        return data    
    
    def run_door_average(df_name):
        new_df = df_name.dropna(axis = 0, how='any')
        st.write(new_df)
        new_df = split_str(new_df)
        HT_average = D_Average_HTS(new_df)
        st.write("Running Door HT Average")
        HT_Name = pd.DataFrame()
        HT_Name = HT_average[HT_average['Fragrance']==name]
        st. write(HT_Name)
        fig= px.bar(HT_Name, x = 'Door', y= 'Average_HTS',title= 'HT Scores by Door')
        st.plotly_chart(fig)
        return fig,new_df,HT_average
    
    def run_accuracy_average(df_name):
        new_df = df_name.dropna(axis = 0, how='any')
        st.write(new_df)
        new_df = split_str(new_df)
        HT_average = Accurate_Average(new_df,name)
        st.write("Running fragrance Accuracy average")
        HT_Name = pd.DataFrame()
        HT_Name = HT_average
        st. write(HT_Name)
        fig= px.bar(HT_Name, x = 'Door', y= 'Average_Accuracy',title= 'Accuracy Scores by Door')
        st.plotly_chart(fig)
        return fig,new_df,HT_average
    
    def Accurate_Average(df_name,name):
        # Filter data where 'Fragrance' equals the specified name
        filtered_df = df_name[df_name['Fragrance'] == name]

        # Group the filtered data by 'Door', 'Fragrance', and 'Agree or Disagree'
        grouped_df = filtered_df.groupby(['Door', 'Fragrance'])

        # Calculate the average 'Agree or Disagree' for each grouping
        average_accuracy = grouped_df['Agree or Disagree'].mean().round(2)

        # Create a new DataFrame with the results
        data = pd.DataFrame({
            'Door': average_accuracy.index.get_level_values(0),
            'Fragrance': average_accuracy.index.get_level_values(1),
            'Average_Accuracy': average_accuracy})
        def reformat(df_name):
            #rename columns, then drop them and reset index to clean up df
            df_name = df_name.rename(columns={'Fragrance': 'Frag', 'Door': 'Door'})
            df_name = df_name.drop(['Door', 'Frag'], axis=1)
            df_name = df_name.reset_index()
        return df_name

        data = reformat(data)
        return data
    
    def run_frag_accuracy_average(df_name):
        new_df = df_name.dropna(axis = 0, how='any')
        st.write(new_df)
        new_df = split_str(new_df)
        HT_average = all_frag_Accurate_Average(new_df,name)
        HT_average = HT_average[HT_average['Fragrance'] != "Fragrance"]
        st.write("Running fragrance Accuracy average")
        HT_Name = pd.DataFrame()
        HT_Name = HT_average
        st. write(HT_Name)
        fig= px.bar(HT_Name, x = 'Fragrance', y= 'Average_Accuracy', color='Door',title= 'Accuracy Scores by Door for all Fragrances',barmode='group')
        st.plotly_chart(fig)
        return fig,new_df,agree_average
    
    def all_frag_Accurate_Average(df_name,name):
        # Filter data where 'Fragrance' equals the specified name
        filtered_df = df_name

        # Group the filtered data by 'Door', 'Fragrance', and 'Agree or Disagree'
        grouped_df = filtered_df.groupby(['Door', 'Fragrance'])

        # Calculate the average 'Agree or Disagree' for each grouping
        average_accuracy = grouped_df['Agree or Disagree'].mean()

        # Create a new DataFrame with the results
        data = pd.DataFrame({
            'Door': average_accuracy.index.get_level_values(0),
            'Fragrance': average_accuracy.index.get_level_values(1),
            'Average_Accuracy': average_accuracy})
        def reformat(df_name):
        #rename columns, then drop them and reset index to clean up df
            df_name = df_name.rename(columns={'Fragrance': 'Frag', 'Door': 'Door'})
            df_name = df_name.drop(['Door', 'Frag'], axis=1)
            df_name = df_name.reset_index()
        return df_name

        data = reformat(data)
        return data
    
    #run the function
    with header_container:
        st.header('Hot Throw Testing',divider="red")
    # Create a file uploader widget
    excel = pd.DataFrame()
    # Check if a file was uploaded
    if uploaded_file is not None:
        # Read the uploaded file into a pandas DataFrame
        excel = pd.read_excel(uploaded_file,dtype=str,skiprows=2)
        excel_original = pd.read_excel(uploaded_file,dtype=str,skiprows=2)
    # Check if the required columns exist in the DataFrame
        optional_columns = ['Like/Dislike?', 'Door','Agree or Disagree', 'How Similar to Control']
        required_columns = ['Name','HTS', 'Fragrance']
        agree_columns = ['Agree or Disagree']
        door_columns = ['Door']
        ##Formatting Dataframe 
        # Rename the 'Hot Throw Score' column to 'HTS', keep one excel original for future use
        excel = excel.rename(columns={'Hot Throw Score': 'HTS'})
        excel_original = excel_original.rename(columns={'Hot Throw Score': 'HTS'})
        excel_original = excel_original.dropna(how='all')
        #create list of all unique fragrances in DF
        unique_name = excel['Fragrance'].unique()
        #Give User options box to select fragrance of interest
        with header_container:
            name = st.selectbox("Name of Fragrance to Analyze", unique_name)
        #New DF for modifying data for HT tests,modifying hyphonated data.
        New_excel = pd.DataFrame()
        try:
            spl_excel = split_str(excel_original)
            spl_excel['Fragrance'] = spl_excel['Fragrance'].astype(str)
            st.write(spl_excel)
            pass
        except:
            pass
        fig = None
        ##
        with header_container:
            #apply functions to dataframes
            if st.button("Display Monday.com Data"):
                st.header("Monday.com Data")
                st.write(excel)
            #create containers to organize elements of called functions
        HTS_container = st.container(border=True)
        NF_container = st.container(border=True)
        Report_container = st.container(border=True)
        
        with header_container:
            #Choosing what option to test for
            st.header("What Type of Test is this?")
            #buttons to determine what functions are done
            if st.checkbox("Hot Throw Analysis",key="HTA_key"):

                with HTS_container:
                    st.title("Run Hot Throw Analysis")
                    #Buttons to run functions for Personal and Fragrance Hot throw
                    st.header("How many HT rooms were used in this test?")
                    doors = st.select_slider("Number of HT Rooms used",options=(1,2,3),key="HT_room")
                    if doors == 1 and set(required_columns).issubset(excel.columns):
                        if st.button(f"Hot Throw Average for {name}"):
                            with Report_container:
                                st.title("Report")
                                HT_df = excel[required_columns]
                                fig,new_df,Avg_HTS,filtered_df = Average_HTS(HT_df,name)
                                st.markdown(f"<p style='color: red; font-size: 24px; '>Hot Throw Average for, {name} is {Avg_HTS} </p>",unsafe_allow_html=True)
                                report = create_excel_report_with_plotly_fig_and_chart(fig,None,new_df, filtered_df,f"Hot Throw Average for{name} report.xlsx")
                                report_path = report
                                st.session_state.reports[f"Hot Throw Average for{name} report.xlsx"]=report_path
                                # Open the file in binary mode for download
                                with open(report, "rb") as file:
                                    st.download_button('Download Report', data=file, file_name=f"Hot Throw Average for {name} report.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                        if st.button("Hot Throw Average for all Fragrances"):
                            with Report_container:
                                st.title("Report")
                                HT_df = excel[required_columns]
                                fig, new_df, HT_average = run_NHT_average(HT_df)
                                report = create_excel_report_with_plotly_fig_and_chart(fig,None,new_df, HT_average,'Hot Throw Average for all Fragrances report.xlsx')
                                report_path = report
                                st.session_state.reports['Hot Throw Average for all Fragrances report.xlsx']=report_path
                                # Open the file in binary mode for download
                                with open(report, "rb") as file:
                                    st.download_button('Download Report', data=file, file_name='all_frag_report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                                                                
                    if doors == 1 and set(required_columns).issubset(excel.columns) and set(agree_columns).issubset(excel.columns):
                        if st.button(f"Agree/Disagree Average for {name}"):
                            with Report_container:
                                st.title("Report")
                                ap_df = excel[required_columns + ['Agree or Disagree']]
                                fig, new_df,Ap_AVG,filtered_df = Ap_avg(ap_df,name)
                                st.markdown(f"<p style='color: red; font-size: 24px;'>The Average Liklienss for {name} is {Ap_AVG}</p>",unsafe_allow_html=True)
                                report = create_excel_report_with_plotly_fig_and_chart(fig,None,new_df, filtered_df,f"The Average Liklienss for {name}report.xlsx")
                                report_path = report
                                st.session_state.reports[f"The Average Liklienss for {name}report.xlsx"]=report_path
                                # Open the file in binary mode for download
                                with open(report, "rb") as file:
                                    st.download_button('Download Report', data=file, file_name=f"Agree_{name}_report.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    if st.button("Agree/Disagree Average for all Fragrances"):
                        with Report_container:
                                st.title("Report")
                                door_df = excel[required_columns + ['Agree or Disagree']]
                                fig, new_df, agree_average = run_agree_avg(door_df)
                                report = create_excel_report_with_plotly_fig_and_chart(fig,None,new_df, agree_average,'Agree Average for all fragrances report.xlsx')
                                report_path = report
                                st.session_state.reports['Agree Average for all fragrances report.xlsx']=report_path
                                # Open the file in binary mode for download
                                with open(report, "rb") as file:
                                    st.download_button('Download Report', data=file, file_name='report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')                 
                #Buttons to run functions for multiple Doors
                    if doors != 1 and set(required_columns).issubset(excel.columns)and set(door_columns).issubset(excel.columns):
                        if st.button(f"Door Hot Throw Average for {name}"):
                            with Report_container:
                                st.title("Report")
                                door_df = excel[required_columns + ['Door']]
                                fig,new_df,filtered_df= run_door_average(door_df)
                                report = create_excel_report_with_plotly_fig_and_chart(fig,None,new_df, filtered_df,f"Door Hot Throw Average for {name} report.xlsx")
                                report_path = report
                                st.session_state.reports[f"Door Hot Throw Average for {name} report.xlsx"]=report_path        
                                # Open the file in binary mode for download
                                with open(report, "rb") as file:
                                    st.download_button('Download Report', data=file, file_name=f"Door Hot Throw Average for {name} report.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                        if st.button("All Fragrances Door Hot Throw Average"):
                            with Report_container:
                                st.title("Report")
                                door_df = excel[required_columns + ['Door']]
                                fig,new_df,filtered_df = run_frag_door_average(door_df)
                                report = create_excel_report_with_plotly_fig_and_chart(fig,None,new_df, filtered_df,'All Fragrance Door Hot Throw Average report.xlsx')
                                report_path = report
                                st.session_state.reports['All Fragrance Door Hot Throw Average report.xlsx']=report_path 
                                # Open the file in binary mode for download
                                with open(report, "rb") as file:
                                    st.download_button('Download Report', data=file, file_name='all_frag_door_report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                        if st.button(f"Door Agree/Disagree Average for {name}"):
                            with Report_container:
                                st.title("Report")
                                door_df = excel[required_columns + ['Agree or Disagree','Door']]
                                fig,new_df,filtered_df = run_accuracy_average(door_df)
                                report = create_excel_report_with_plotly_fig_and_chart(fig,None,new_df, filtered_df,f"Door Agree/Disagree Average for {name} report.xlsx")
                                report_path = report
                                st.session_state.reports[f"Door Agree/Disagree Average for {name} report.xlsx"]=report_path 
                                # Open the file in binary mode for download
                                with open(report, "rb") as file:
                                    st.download_button('Download Report', data=file, file_name=f"Door Agree/Disagree Average for {name}'report.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

                        if st.button("All Fragrances Door Agree/Disagree Average"):
                            with Report_container:
                                st.title("Report")
                                door_df = excel[required_columns + ['Agree or Disagree','Door']]
                                fig,new_df,agree_average = run_frag_accuracy_average(door_df)
                                report = create_excel_report_with_plotly_fig_and_chart(fig,None,new_df, agree_average,'All Fragrance Door Agree Average report.xlsx')
                                report_path = report
                                st.session_state.reports['All Fragrance Door Agree Average report.xlsx']=report_path 
                                # Open the file in binary mode for download
                                with open(report, "rb") as file:
                                    st.download_button('Download Report', data=file, file_name='report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')                         
                                        



    
#initialize session state for reports
if 'reports' not in st.session_state:
    st.session_state.reports = {}
#title to main page
with title_container:
    st.title("Aryballe Data Analysis")
#sidebar title
st.sidebar.title("Select a function")
#first checkbox for sensory panel upload
if st.sidebar.checkbox('Sensory Panel Data Upload',key="sensory"):
    
    with header_container:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
        
    sensory_panel(uploaded_file)
#checkbox to run Aryballe Data upload
if st.sidebar.checkbox("Aryballe Data",key="Aryballe"):
    #UI setup and data handling
    with header_container:
        st.title('Aryballe Data')
        #Create a file uploader widget
        uploaded_files = st.file_uploader("Choose and excel file",type = "csv",accept_multiple_files=True)
        #create session state dictionary to store dataframes
        if 'dataframes' not in st.session_state:
            st.session_state['dataframes'] = {}
            st.session_state['intensity'] = {}

        excel = pd.DataFrame()
        start_time = float(1)
        #Check if file was uploaded


        if uploaded_files is not None:
            if st.button("Upload"):
                #clear session_state from previous uploads
                st.session_state['dataframes'] = {}

                for uploaded_file in uploaded_files:#clear session_state for 'dictionary'
                    
                    #Read the uploaded file and save to pd dataframe
                    temp_df = pd.read_csv(uploaded_file,header= None, dtype=str, sep=';')
                    #Get headers from first row, split by ;
                    headers = temp_df.iloc[0]
                    #drop header row from df
                    temp_df = temp_df.drop(0)
                    #set columns to headers
                    temp_df.columns = headers
                    # Reset the index of the DataFrame
                    temp_df.reset_index(drop=True, inplace=True)
                    temp_df['Time'] = convert_unix_to_time_duration(temp_df,'timestamp')
                    st.session_state['dataframes'][uploaded_file.name] = temp_df
                st.write("The Dataframes uploaded are:",list(st.session_state['dataframes'].keys()))
            # Example debug line to check the content of `dataframes`
            st.write(f"Debug: Contents of dataframes dict: {list(st.session_state['dataframes'].keys())}")


            if st.button("View File"):
                for filename, df in st.session_state['dataframes'].items():
                            st.write(f"Dataframe for {filename}")
                            st.write(df)


            if st.checkbox("Intensity Calcultion by Cycle"):
                col_type = 'cycle'
                with report_container:
                    st.title("Report")
                    spot_col = ['spot22','spot23','spot24','spot25','spot28','spot29','spot65','spot66']
                    for filename, df in st.session_state['dataframes'].items():
                        mean_intensity= df[spot_col].astype(float).mean(axis=1)
                        temp_df = pd.DataFrame({'run_id':df['run-id'],'cycle': df['cycle'], 'mean intensity': mean_intensity})
                        st.session_state['intensity'][filename] = temp_df
                    for filename, df in st.session_state['intensity'].items():
                            st.header(f"Dataframe for {filename}",divider="red")
                            # Plot using Plotly Express, faceting by 'run_id'

                            fig = create_bar(df,'cycle','mean intensity',f"Intensity for {filename}",'cycle',None)

                            if st.checkbox("Average Intensity within time range",key=filename):
                                                
                                    avg, avg_fig = avg_intensity_time_range(df,col_type,filename)
                                    report = create_excel_report_with_plotly_fig_and_chart(fig,avg_fig,df,avg,f'{filename} report.xlsx')
                                    report_path = report
                                    st.session_state.reports[f'{filename} report.xlsx']=report_path
                                    # Open the file in binary mode for download
                                    with open(report, "rb") as file:
                                        st.download_button('Download Report', data=file, file_name=f'{filename} report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')        
                                    st.divider()
            if st.checkbox("Intensity Calculation by Time"):
                col_type = 'Time'
                with report_container:
                    st.title("Report")
                    st.session_state['intensity'] = {}
                    Time = []
                    spot_col = ['spot22','spot23','spot24','spot25','spot28','spot29','spot65','spot66']
                    for filename, df in st.session_state['dataframes'].items():
                        mean_intensity= df[spot_col].astype(float).mean(axis=1)
                        temp_df = pd.DataFrame({'run_id':df['run-id'],'cycle': df['cycle'], 'mean intensity': mean_intensity,'Time':df['Time']})
                        st.session_state['intensity'][filename] = temp_df
                    for filename, df in st.session_state['intensity'].items():
                            st.header(f"Dataframe for {filename}",divider="red")
                            # Plot using Plotly Express, faceting by 'run_id'
                            fig = create_bar(df,'Time','mean intensity',f"Intensity for {filename}",'Time',None)

                            if st.checkbox("Average Intensity within time range",key=filename):           
                                    avg, avg_fig = avg_intensity_time_range(df,col_type,filename)
                                    report = create_excel_report_with_plotly_fig_and_chart(fig,avg_fig,df,avg,f'{filename} report.xlsx')
                                    report_path = report
                                    st.session_state.reports[f'{filename} report.xlsx']=report_path
                                    # Open the file in binary mode for download
                                    with open(report, "rb") as file:
                                        st.download_button('Download Report', data=file, file_name=f'{filename} report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')        

        else:
            st.write("No file uploaded")

if st.session_state.reports:
    selected_reports = st.sidebar.multiselect("Select reports to compile", options=list(st.session_state.reports.keys()), default=None)
    if st.sidebar.button ("Combine Reports you selected?"):
        if selected_reports:
            #Combine the reports
            output_filename = combine_excel_sheets(selected_reports)
            #provide download link
            with open(output_filename,"rb") as file:
                st.sidebar.download_button("Download combined Report",file,file_name=output_filename)
        
    else:
        st.sidebar.write("No reports available.")