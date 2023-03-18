
# Import libraries
import bcrypt
import streamlit as st
from enrich import showEnrichPage


# ====================================================================================================


# Set tab info
st.set_page_config(
    page_title="2X | Prospiant",
    page_icon="🎏"
)


# ====================================================================================================


# Initialize session variables

# Variable to check if user is logged in
if 'login_status' not in st.session_state:
    st.session_state['login_status'] = False

# Variable to check if attempt was made
if 'failed_attempt' not in st.session_state:
    st.session_state['failed_attempt'] = False


# ====================================================================================================


# Function to check if code matches
def checkCode() :

    # Set real password's hash
    hash_list = [
        b'$2b$12$ZeHtfYRuDZMwjkCbJWwe7enHtBAOAUnK1O2w/ZxLEK9w6to0TEVhm',
        b'$2b$12$HcmMrhPEfp1d7lBzRslPUuL3qERDaS/5TxlG/MNAhaKIKsTYQzBxS'
    ]

    # Encode the input text
    encoded_text_input = st.session_state['text_input'].encode('utf-8')

    # Initialize match variable
    matched = False

    # Loop through the hash list
    for hash in hash_list :

        # Compare encoded input text with hash
        matched = bcrypt.checkpw(encoded_text_input, hash)

        # When found a match
        if matched :

            # Stop loop
            break

    # Set path for post checking action
    if matched :

        # Change login status for matched
        st.session_state['login_status'] = True
    
    else : 

        # Clear the input text box
        st.session_state['text_input'] = ''

        # Keep login status as before
        st.session_state['login_status'] = False

        # Trigger failed attempt 
        st.session_state['failed_attempt'] = True


# Function to display login page
def showLoginPage() :

    # Set the content alignment
    left_space, content, right_space = st.columns([1, 3, 1])

    # Define content
    with content :
    
        # Set logo
        st.image('prospiant-logo.png', use_column_width = 'always')

        # Set the inner content alignment
        inner_left_space, inner_content, inner_right_space = st.columns([1, 3, 1])
        
        # Define inner content
        with inner_content:

            # Create space
            st.write('')

            # Set input text box
            text_input = st.text_input(
                label = 'Enter code',
                key = 'text_input', 
                type = 'password',
                autocomplete = None,
                on_change = checkCode,
                placeholder = 'Type here',
                label_visibility = 'collapsed'
            )

            # Show message below input text box
            if not st.session_state['failed_attempt'] :

                # Display guide message
                st.info('Enter access code above.')

            else :
                
                # Display error message
                st.error('Invalid code, try again.')


    # Vertically centre the content of the page 
    st.write(
        """
        <style>
        [data-testid="stAppViewContainer"] {
            padding: 10vh 1vw;
        }
        </style>
        """,
        unsafe_allow_html=True
    )


# ====================================================================================================


# Run main workflow of the page

# Check login status
if not st.session_state['login_status'] :

    # Display login page when not logged in
    showLoginPage()

else :

    # Display navigation bar when logged in
    showEnrichPage()
    
    
# Hide the hamburger icon and "Made with Streamlit" footer
st.write(
    """
    <style>
    #MainMenu {
        visibility: hidden;
    }
    footer {
        visibility: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True
)

