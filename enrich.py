
# Import libraries
import streamlit as st
import numpy as np
import pandas as pd
from random import randint


# Function to display enrich sheets page
def showEnrichPage() :


    # Adjust length of button
    st.write(
        """
        <style>
        [class="row-widget stButton"] button {
            width: 100%;
            background-color: #FD6767;
        }
        [class="row-widget stDownloadButton"] button {
            width: 50%;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # To store file uploader key
    if 'file_upload' not in st.session_state:
        st.session_state['file_upload'] = str(randint(1000, 100000000))

    # To store process ended boolean
    if 'process_ended' not in st.session_state:
        st.session_state['process_ended'] = True
    
    # To store output table
    if 'output_table' not in st.session_state:
        st.session_state['output_table'] = []

    # To store delivered dataframe
    if 'delivered_data' not in st.session_state:
        st.session_state['delivered_data'] = []

    # To store opened dataframe
    if 'opened_data' not in st.session_state:
        st.session_state['opened_data'] = []
    
    # To store clicked dataframe
    if 'clicked_data' not in st.session_state:
        st.session_state['clicked_data'] = []


    # =====================================================================================================================


    # Set page title
    st.title('Enrich Sheets :card_file_box:')

    # Create space
    st.write('')

    # First step -- Upload Excel files
    st.subheader(':inbox_tray: Upload Excel Files')
    
    # Create space
    st.write('')
    
    # Create file uploader for Excel files
    uploaded_files = st.file_uploader(
        label = 'Upload Excel files', 
        type = ['xls', 'xlsx'],
        accept_multiple_files = True,
        key = st.session_state['file_upload'],
        on_change = resetProcess,
        label_visibility = 'collapsed'
    )
    
    # When there are no uploaded files
    if not uploaded_files :

        # Set process to end
        st.session_state['process_ended'] = True

        # Clear out stored data
        st.session_state['output_table'] = []
        st.session_state['delivered_data'] = []
        st.session_state['opened_data'] = []
        st.session_state['clicked_data'] = []


    # Create space
    st.write('')
    

    # =====================================================================================================================


    # Second step -- Start Enrichment
    st.subheader(':postal_horn: Start Process')
    
    # Create space
    st.write('')

    # Create button for starting process
    start_button = st.button(
        label = '**Enrich All Excel Files**',
        type = 'primary',
        on_click = endProcess,
        disabled = st.session_state['process_ended']
    )

    # When button is pressed
    if start_button :

        # Create spinner for loading
        with st.spinner('In progress...'):

            # Enrich data
            enrichSheets(uploaded_files)
    

    # When there is output
    if len(st.session_state['output_table']) > 0 :

        # Set label
        st.write('List of campaigns found:')

        # Display dataframe
        st.dataframe(st.session_state['output_table'], use_container_width = True)

        # Set label
        st.write('Export campaigns data:')

        # When there is data for delivered
        if len(st.session_state['delivered_data']) > 0 :

            # Display download button
            st.download_button(
                label = 'Download All Campaign Delivered Data',
                data = st.session_state['delivered_data'].to_csv(encoding = 'utf-8-sig', index = False),
                file_name = 'Compiled Delivered Data.csv',
                mime = 'text/csv'
            )

        # When there is data for opened
        if len(st.session_state['opened_data']) > 0 :

            # Display download button
            st.download_button(
                label = 'Download All Campaign Open Data',
                data = st.session_state['opened_data'].to_csv(encoding = 'utf-8-sig', index = False),
                file_name = 'Compiled Opened Data.csv',
                mime = 'text/csv'
            )

        # When there is data for clicked
        if len(st.session_state['clicked_data']) > 0 :

            # Display download button
            st.download_button(
                label = 'Download All Campaign Click Data',
                data = st.session_state['clicked_data'].to_csv(encoding = 'utf-8-sig', index = False),
                file_name = 'Compiled Clicked Data.csv',
                mime = 'text/csv'
            )

        # Create space
        st.write('')

        # Create checkbox for restart
        if st.checkbox('Restart') :

            # Change the key of the file uploader
            st.session_state['file_upload'] = str(randint(1000, 100000000))

            # Set process to end
            st.session_state['process_ended'] = True

            # Clear out stored data
            st.session_state['output_table'] = []
            st.session_state['delivered_data'] = []
            st.session_state['opened_data'] = []
            st.session_state['clicked_data'] = []

            # Rerun the page
            st.experimental_rerun()


# Function to reset process - for file uploader
def resetProcess() :

    # Disable pressed start trigger
    st.session_state['process_ended'] = False

    # Clear out stored data
    st.session_state['output_table'] = []
    st.session_state['delivered_data'] = []
    st.session_state['opened_data'] = []
    st.session_state['clicked_data'] = []


# Function to end process - for start button
def endProcess() :

    # Enable pressed start trigger
    st.session_state['process_ended'] = True


# Function to enrich data
def enrichSheets(uploaded_files) :

    # Initialize lists
    campaign_list = []
    delivered_count_list = []
    opened_count_list = []
    clicked_count_list = []


    # Loop through each file
    for file in uploaded_files :

        # Try block for reading data in
        try : 

            # Read data in from Report Summary sheet
            df_summary = pd.read_excel(file, sheet_name = "Report Summary", header = None)

            # Read data in from Campaign Delivery sheey
            df_delivered = pd.read_excel(file, sheet_name = "Campaign Delivery")

            # Read data in from Campaign Delivery sheey
            df_opened = pd.read_excel(file, sheet_name = "Opens")

            # Read data in from Campaign Delivery sheey
            df_clicked = pd.read_excel(file, sheet_name = "Clicks")

        # When there is error
        except : 

            # Display error message
            st.error('Error in getting data from the sheets in *{x}* file.'.format(x = file.name))

            # Skip the file
            continue

        # =======================================================================================================================
        # =======================================================================================================================

        # Try block for getting campaign name
        try :
            
            # Obtain campaign name
            campaign_name = df_summary.iloc[0, 1]

            # Append campaign name to list
            campaign_list.append(campaign_name)

        # When there is error
        except :
            
            # Display error message
            st.error('Error in finding the campaign name from the sheets in *{x}* file.'.format(x = file.name))

            # Skip the file
            continue

        # =======================================================================================================================
        # =======================================================================================================================

        # Try block for inserting campaign name and removing zeros
        try :
            
            # Delivered

            # Add campaign name to delivered data
            df_delivered['Campaign Name'] = campaign_name

            # Remove zeros from delivered data
            df_delivered = df_delivered.replace(['0'], np.nan)

            # Get total rows
            delivered_count_list.append(len(df_delivered))

            # ==================================================================

            # Opened

            # Add campaign name to opened data
            df_opened['Campaign Name'] = campaign_name

            # Remove zeros from opened data
            df_opened = df_opened.replace(['0'], np.nan)

            # Get total rows
            opened_count_list.append(len(df_opened))


            # ==================================================================

            # Clicked

            # Add campaign name to clicked data
            df_clicked['Campaign Name'] = campaign_name

            # Remove zeros from clicked data
            df_clicked = df_clicked.replace(['0'], np.nan)

            # Get total rows
            clicked_count_list.append(len(df_clicked))

        # When there is error
        except :
            
            # Display error message
            st.error('Error in inserting the campaign name into main data for *{x}* file.'.format(x = file.name))

            # Skip the file
            continue
        
        # =======================================================================================================================
        # =======================================================================================================================

        # Try blocking for appending data
        try :

            # When there is no data for delivered
            if len(st.session_state['delivered_data']) == 0 :

                # Assign the first dataframe
                st.session_state['delivered_data'] = df_delivered

            # When there is data for delivered
            else :

                # Get existing dataframe
                old_df = st.session_state['delivered_data']

                # Append new dataframe to the old dataframe
                st.session_state['delivered_data'] = old_df.append(df_delivered, ignore_index = True)
            
            # ==================================================================================================

            # When there is no data for opened
            if len(st.session_state['opened_data']) == 0 :

                # Assign the first dataframe
                st.session_state['opened_data'] = df_opened

            # When there is data for opened
            else :

                # Get existing dataframe
                old_df = st.session_state['opened_data']

                # Append new dataframe to the old dataframe
                st.session_state['opened_data'] = old_df.append(df_opened, ignore_index = True)

            # ==================================================================================================
            
            # When there is no data for clicked
            if len(st.session_state['clicked_data']) == 0 :

                # Assign the first dataframe
                st.session_state['clicked_data'] = df_clicked

            # When there is data for clicked
            else :

                # Get existing dataframe
                old_df = st.session_state['clicked_data']

                # Append new dataframe to the old dataframe
                st.session_state['clicked_data'] = old_df.append(df_clicked, ignore_index = True)

        # When there is error
        except :

            # Display error message
            st.error('Error in combining data from *{x}* file with the data from previous files.'.format(x = file.name))

            # Skip the file
            continue

    # ========================================================

    # Create output dictionary
    dict_of_lists = {
        'Campaign Name': campaign_list,
        'Delivered': delivered_count_list,
        'Open': opened_count_list,
        'Click': clicked_count_list,
    }

    # Create output dataframe
    output_table = pd.DataFrame(dict_of_lists)

    # Start table from index 1
    output_table.index += 1

    # Set session output table
    st.session_state['output_table'] = output_table

     