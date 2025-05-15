import json
import logging
import os
import shutil
import sys
import tempfile
from pathlib import Path
import streamlit as st
import yaml
import base64

# --- Constants ---
PAGE_TITLE = "PDF to PowerPoint Generator"
APP_TITLE = "SlidesGen"
PAGE_PRESENTATION_GENERATOR = "Presentation Generator"
PAGE_CONFIGURATION_EDITOR = "Configuration Editor"
PAGES = [PAGE_PRESENTATION_GENERATOR, PAGE_CONFIGURATION_EDITOR]

# --- Set up logging ---
logging.basicConfig(stream=sys.stdout, level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
log = logging.getLogger(__name__)
# Prevent duplicate log handlers if script reruns
if not log.handlers:
    log.addHandler(logging.StreamHandler(stream=sys.stdout))
    log.propagate = False

# --- Set page config and custom CSS ---
st.set_page_config(page_title=PAGE_TITLE, layout="wide")

# Apply custom CSS for background colors
st.markdown("""
    <style>
        .stApp {
            background: linear-gradient(to bottom, #77B1D4, #517891);
        }
        /* Style the top toolbar/header bar */
        header[data-testid="stHeader"] {
            background-color: #607C8E !important;
        }
        /* Style the hamburger menu icon and buttons in header */
        header[data-testid="stHeader"] button {
            background-color: transparent !important;
            color: black !important;
        }
        /* Style the fullscreen button and other elements in the top bar */
        [data-testid="stToolbar"] {
            background-color: #607C8E !important;
            right: 0;
        }
        [data-testid="stToolbar"] button {
            color: black !important;
        }
        /* Additional styling for top bar elements */
        [data-testid="stDecoration"] {
            background-color: #607C8E !important;
        }
        div[data-testid="stStatusWidget"] {
            background-color: #607C8E !important;
            color: black !important;
        }
        #MainMenu {
            background-color: #607C8E !important;
        }
        .stSidebar {
            background-color: rgba(255, 255, 255, 0.1);
        }
        .stMarkdown {
            color: black;
        }
        .stTitle, .stHeader, h1, h2, h3 {
            color: black !important;
        }
        div[data-testid="stText"] {
            color: black;
        }
        label, .stRadio label, .stSelectbox label, .stFileUploader label, .stTextInput label {
            color: black !important;
        }
        /* Grey buttons */
        div.stButton > button {
            background-color: #607C8E !important;
            color: white !important;
            border-color: #4A6175 !important;
        }
        div.stButton > button:hover {
            background-color: #4A6175 !important;
            color: white !important;
            border-color: #35485C !important;
        }
        /* Grey dropdowns */
        div.stSelectbox > div > div > div {
            background-color: #607C8E !important;
            color: white !important;
        }
        .stSelectbox [data-baseweb="select"] {
            background-color: #607C8E !important;
        }
        .stSelectbox [data-baseweb="select"] > div {
            background-color: #607C8E !important;
            color: black !important;
        }
        .stSelectbox [data-baseweb="select"] svg {
            color: white !important;
        }
        .stSelectbox [data-baseweb="select"] div[data-testid="stMarkdown"] p {
            color: white !important;
        }
        /* Grey dropdowns in the sidebar */
        .stSidebar .stSelectbox [data-baseweb="select"] {
            background-color: #607C8E !important;
        }
        .stSidebar .stSelectbox [data-baseweb="select"] > div {
            background-color: #607C8E !important;
            color: black !important;
        }
        /* Additional sidebar selectbox styling */
        .stSidebar .stSelectbox div[role="listbox"] {
            background-color: #607C8E !important;
        }
        .stSidebar .stSelectbox div[role="option"] {
            background-color: #607C8E !important;
            color: black !important; 
        }
        .stSidebar .stSelectbox div[role="option"]:hover {
            background-color: #4A6175 !important;
        }
        /* The dropdown arrow */
        .stSidebar .stSelectbox svg {
            color: black !important;
        }
        /* Style for radio buttons */
        .stRadio > div {
            background-color: #607C8E !important;
            color: white !important;
            padding: 10px;
            border-radius: 5px;
        }
        .stRadio label {
            color: white !important;
        }
        /* Additional elements that need styling */
        .stFileUploader > div > div {
            background-color: #607C8E !important;
        }
        .stFileUploader span {
            color: black !important;
        }
        /* File uploader drag area styling */
        [data-testid="stFileUploader"] section {
            background-color: #607C8E !important;
            border-color: rgba(0, 0, 0, 0.2) !important;
        }
        [data-testid="stFileUploader"] section p {
            color: black !important;
        }
        [data-testid="stFileUploader"] section small {
            color: black !important;
        }
        /* Text input styling */
        .stTextInput > div > div > input {
            background-color: #607C8E !important;
            color: black !important;
            border-color: rgba(0, 0, 0, 0.2) !important;
        }
        /* Browse files button within uploader */
        [data-testid="stFileUploader"] section button {
            background-color: #4A6175 !important;
            color: black !important;
            border-color: rgba(0, 0, 0, 0.2) !important;
        }
        /* Additional styling for menu items and dropdowns */
        div[data-baseweb="popover"] div {
            background-color: #607C8E !important;
        }
        div[data-baseweb="popover"] li, div[data-baseweb="popover"] a {
            color: black !important;
        }
        div[data-baseweb="popover"] li:hover {
            background-color: #4A6175 !important;
        }
        /* Logo positioning - shift slightly right */
        [data-testid="column"]:nth-of-type(2) [data-testid="stImage"] {
            margin-left: 10%;
        }
    </style>
""", unsafe_allow_html=True)

# --- Attempt to load dotenv ---
try:
    from dotenv import load_dotenv
    load_dotenv()
    log.info("Environment variables loaded from .env.")
except ImportError:
    log.warning("python-dotenv not installed. Relying solely on system environment variables or direct input.")
except Exception as e:
    log.warning("Error loading environment variables from .env: %s", e)


# --- LlamaIndex Imports ---
try:
    from llama_index.core import (
        Settings, SimpleDirectoryReader)
    from llama_index.embeddings.google_genai import GoogleGenAIEmbedding
    from llama_index.llms.google_genai import GoogleGenAI
    log.info("LlamaIndex core components and Google models imported.")
except ImportError as e:
    st.error(f"Required LlamaIndex components or Google models not found: {e}")
    st.error("Please ensure you have installed the necessary packages, e.g., `pip install llama-index-core llama-index-llms-google-genai llama-index-embeddings-google-genai python-pptx PyYAML`")
    st.stop()
except Exception as e:
    st.error(f"An unexpected error occurred during LlamaIndex import: {e}")
    st.stop()

# --- Local Module Imports ---
try:
    import config
    import indexing
    import presentation
    import utils
    log.info("Local modules (config, utils, indexing, presentation) imported.")
except ImportError as e:
    st.error(f"Failed to import local modules (config.py, utils.py, indexing.py, presentation.py). Ensure they exist in the same directory: {e}")
    st.stop()
except Exception as e:
    st.error(f"An unexpected error occurred during local module import: {e}")
    st.stop()


# --- Configuration Editor Callbacks ---

def save_config_to_file():
    """Callback to save editor content to the config file."""
    log.info("Callback: Saving config editor content.")
    config_file_path = Path(config.CONFIG_FILE_NAME)
    content_to_save = st.session_state.get('config_editor_content', '')

    try:
        # Attempt to parse the edited content to validate
        yaml.safe_load(content_to_save)
        # If parsing succeeds, write to file
        with open(config_file_path, 'w', encoding='utf-8') as f:
            f.write(content_to_save)
        st.session_state.message = f"Configuration saved successfully to `{config.CONFIG_FILE_NAME}`."
        st.session_state.message_type = "success"
        log.info("Configuration saved successfully in callback.")
        # Update raw_config_text state as well after successful save
        st.session_state.raw_config_text = content_to_save

    except yaml.YAMLError as e:
        st.session_state.message = f"Error saving configuration: Invalid YAML format. Details: {e}"
        st.session_state.message_type = "error"
        log.error("Failed to save config (callback): Invalid YAML format.", exc_info=True)
    except IOError as e:
        st.session_state.message = f"An error occurred while saving the file: {e}"
        st.session_state.message_type = "error"
        log.error("IOError during config save (callback).", exc_info=True)
    except Exception as e:
        st.session_state.message = f"An unexpected error occurred while saving: {e}"
        st.session_state.message_type = "error"
        log.error("Unexpected error during config save (callback).", exc_info=True)

def delete_config_file():
    """Callback to delete the config file."""
    log.info("Callback: Deleting config file: %s.", config.CONFIG_FILE_NAME)
    config_file_path = Path(config.CONFIG_FILE_NAME)
    try:
        config_file_path.unlink(missing_ok=True)
        st.session_state.message = f"Configuration file `{config.CONFIG_FILE_NAME}` deleted (or did not exist)."
        st.session_state.message_type = "warning"
        log.info("Configuration file deleted successfully in callback.")
        # Clear the editor content in state
        st.session_state.raw_config_text = ""
        st.session_state.config_editor_content = ""
    except OSError as e:
        st.session_state.message = f"An error occurred while deleting: {e}"
        st.session_state.message_type = "error"
        log.error("OSError during config delete (callback).", exc_info=True)
    except Exception as e:
        st.session_state.message = f"An error occurred while deleting: {e}"
        st.session_state.message_type = "error"
        log.error("Unexpected error during config delete (callback).", exc_info=True)

def reload_config_from_file():
    """Callback to reload editor content from the config file."""
    log.info("Callback: Reloading config from file: %s", config.CONFIG_FILE_NAME)
    try:
        _, reloaded_raw_text, reloaded_file_exists, parse_successful = utils.load_config_file(config.CONFIG_FILE_NAME)
        if reloaded_file_exists:
            st.session_state.raw_config_text = reloaded_raw_text
            st.session_state.config_editor_content = reloaded_raw_text
            if parse_successful:
                st.session_state.message = f"Content reloaded from `{config.CONFIG_FILE_NAME}`."
                st.session_state.message_type = "success"
            else:
                st.session_state.message = f"Content reloaded from `{config.CONFIG_FILE_NAME}`, but it contains invalid YAML."
                st.session_state.message_type = "warning"
            log.info("Config reloaded in callback (file exists).")
        else:
            st.session_state.raw_config_text = ""
            st.session_state.config_editor_content = ""
            st.session_state.message = f"Configuration file `{config.CONFIG_FILE_NAME}` not found to reload."
            st.session_state.message_type = "info"
            log.info("Config file not found during reload callback.")
    except Exception as e:
        log.error("Error during config reload callback: %s", e, exc_info=True)
        st.session_state.message = f"Error reloading configuration: {e}"
        st.session_state.message_type = "error"

def load_default_config():
    """Callback to load default config content into the editor."""
    log.info("Callback: Loading default config content.")
    st.session_state.config_editor_content = config.DEFAULT_CONFIG_CONTENT
    st.session_state.message = "Default configuration template loaded into editor. Click 'Save Configuration' to create the file."
    st.session_state.message_type = "info"


# --- Streamlit App Layout ---

# Initialize session state for page if not already set
if "page" not in st.session_state:
    st.session_state.page = PAGE_PRESENTATION_GENERATOR

# Initialize session state for indexing options
if "use_persistent_index" not in st.session_state:
    st.session_state.use_persistent_index = False

if "force_rebuild" not in st.session_state:
    st.session_state.force_rebuild = False

# Left Sidebar Content
with st.sidebar:
    st.header("Navigation")
    page = st.radio(
        "Go to",
        PAGES,
        index=PAGES.index(st.session_state.page),
        key='page_radio',
        on_change=lambda: setattr(st.session_state, 'page', st.session_state.page_radio)
    )

    # Presentation/Indexing options only visible on the Generator page
    if st.session_state.page == PAGE_PRESENTATION_GENERATOR:
        st.markdown("---")
        st.header("Presentation Options")

        # --- Load config for theme selection ---
        config_data_info, _, config_file_exists_info, parse_successful_info = utils.load_config_file(config.CONFIG_FILE_NAME)
        available_themes = []
        default_theme_index = 0
        current_theme_from_config = None
        theme_selection_disabled = True
        theme_help_text = f"Create or edit '{config.CONFIG_FILE_NAME}' to manage themes."

        if config_file_exists_info and parse_successful_info and isinstance(config_data_info, dict):
            current_theme_from_config = config_data_info.get('current_theme')
            templates_dict = config_data_info.get('templates')
            if isinstance(templates_dict, dict):
                available_themes = list(templates_dict.keys())
                if available_themes:
                    theme_selection_disabled = False
                    theme_help_text = f"Select a theme defined in '{config.CONFIG_FILE_NAME}'. The template path associated with the theme will be used."
                    if current_theme_from_config in available_themes:
                        try:
                            default_theme_index = available_themes.index(current_theme_from_config)
                        except ValueError:
                            log.warning("current_theme '%s' from config not found in available themes list.", current_theme_from_config)
                            default_theme_index = 0 # Default to first theme if current_theme is invalid
                    else:
                         log.warning("current_theme '%s' not found in templates section.", current_theme_from_config)
                         default_theme_index = 0 # Default to first theme if current_theme is missing or invalid
                else:
                    theme_help_text = f"No themes found in the 'templates' section of '{config.CONFIG_FILE_NAME}'."
            else:
                theme_help_text = f"Config file '{config.CONFIG_FILE_NAME}' is missing the 'templates' dictionary."
        elif not config_file_exists_info:
            theme_help_text = f"Config file '{config.CONFIG_FILE_NAME}' not found. Default styling will be used."
        else: # Not parse_successful_info or not dict
             theme_help_text = f"Config file '{config.CONFIG_FILE_NAME}' is invalid. Default styling will be used."

        # --- Theme Selection Dropdown ---
        selected_theme = st.selectbox(
            "Select Presentation Theme",
            options=available_themes,
            index=default_theme_index,
            key='presentation_theme_selector',
            disabled=theme_selection_disabled,
            help=theme_help_text
        )

        # Add custom CSS to style this specific dropdown
        st.markdown("""
            <style>
                div[data-testid="stSelectbox"] > div:has(label:contains("Select Presentation Theme")) > div > div {
                    background-color: #607C8E !important;
                    color: black !important;
                }
            </style>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.header("Indexing Options")

        # Use on_change handlers to update session state
        def update_persistent():
            st.session_state.use_persistent_index = st.session_state.persistent_checkbox
            
        def update_rebuild():
            st.session_state.force_rebuild = st.session_state.rebuild_checkbox

        use_persistent = st.checkbox(
            f"Use Persistent Index Storage ({config.PERSISTENT_INDEX_DIR})", 
            value=st.session_state.use_persistent_index, 
            key='persistent_checkbox',
            on_change=update_persistent
        )
        st.info(f"If checked, index is saved/loaded from `{config.PERSISTENT_INDEX_DIR}`.")

        force_rebuild = st.checkbox(
            "Force Index Rebuild (if persistent)", 
            value=st.session_state.force_rebuild, 
            disabled=not st.session_state.use_persistent_index, 
            key='rebuild_checkbox',
            on_change=update_rebuild
        )
        st.info("If 'Use Persistent Index Storage' is checked, force rebuilding the index instead of loading.")

# Main layout with columns
main_col, right_col = st.columns([4, 1])

# Main title in main column
with main_col:
    # Display college logo above title with rightward shift
    col1, col2, col3 = st.columns([2, 2, 1])
    with col2:
        # Logo is already shifted through global CSS
        st.image("assets/college_logo.png", width=150)
    
    # Replace standard title with custom markdown to control font size - with reduced spacing
    st.markdown("<h1 style='text-align: center; color: black; font-size: 68px; margin-top: -10px;'>{}</h1>".format(APP_TITLE), unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: black; margin-top: -36px;'>PDF to PPT Generator</h3>", unsafe_allow_html=True)

# Right Sidebar Content - Department Name
with right_col:
    # Logo moved to main column
    st.markdown("<br><h3 style='text-align: center; color: black;'>Department<br>of<br>Computer<br>Engineering</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; color: black;'>Under the guidance of<br>Dr. Chetan Singh Negi</p>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 14px; color: black;'>Developed by:<br>Sheetal Raj<br>Akshita Bhatt<br>Vanshika Arora<br>Kanika Pathak</p>", unsafe_allow_html=True)
    st.markdown("---")

# --- Initialize session state ---
if 'app_state' not in st.session_state:
    st.session_state.app_state = 'initial'
    st.session_state.temp_dir = None
    st.session_state.output_file_path = None

# --- Load config content on initial load for the editor ---
if 'raw_config_text' not in st.session_state:
    log.info("Initial load: Reading config file for editor state.")
    _, initial_raw_text, file_exists, _ = utils.load_config_file(config.CONFIG_FILE_NAME)
    st.session_state.raw_config_text = initial_raw_text if file_exists else ""

# --- Load environment variables for core settings ---
api_key = os.getenv("GEMINI_API_KEY", "")
llm_model = os.getenv("LLM_MODEL", "gemini-1.5-flash-latest")
embedding_model = os.getenv("EMBEDDING_MODEL", "models/embedding-001")

if not api_key and st.session_state.app_state != 'initial':
    st.error("GEMINI_API_KEY environment variable is not set. Please set it in your .env file.")
    st.stop()

# --- Main Content Area ---
if st.session_state.page == PAGE_PRESENTATION_GENERATOR:
    # --- Presentation Generator Page Content ---
    if st.session_state.app_state == 'initial':
        st.write("Upload a PDF, select options, and generate a presentation as a downloadable PowerPoint file.")
        st.markdown("---")

        st.header("Document Input")
        uploaded_pdf = st.file_uploader("Upload your PDF document", type=["pdf"], key="generator_upload")
        output_filename = st.text_input("Output PowerPoint Filename", value="generated_presentation.pptx", key="generator_output_filename")
        if not output_filename.lower().endswith(".pptx"):
            output_filename += ".pptx"

        st.markdown("---")
        generate_button = st.button("Generate Presentation", use_container_width=True, type="primary")

        # --- Logic when Generate Button is clicked ---
        if generate_button:
            # --- Initial Validation ---
            current_api_key = api_key
            current_llm_model = llm_model
            current_embedding_model = embedding_model
            current_selected_theme = selected_theme
            current_use_persistent_index = st.session_state.use_persistent_index
            current_force_rebuild = st.session_state.force_rebuild

            if not current_api_key:
                st.error("API Key is required to proceed.")
                st.session_state.app_state = 'initial'
            elif not uploaded_pdf:
                st.error("Please upload a PDF document.")
                st.session_state.app_state = 'initial'
            else:
                # Validation passed, prepare for Processing
                temp_dir = None
                try:
                    temp_dir = tempfile.mkdtemp()
                    log.info("Created temporary directory: %s", temp_dir)
                    pdf_path_in_temp = utils.save_uploaded_file(uploaded_pdf, temp_dir, filename=uploaded_pdf.name)
                    if not pdf_path_in_temp or not Path(pdf_path_in_temp).exists():
                        raise FileNotFoundError(f"Failed to save uploaded PDF to {temp_dir}")
                    log.info("Uploaded PDF saved to: %s", pdf_path_in_temp)

                    # Store necessary variables in session state
                    st.session_state.temp_dir = temp_dir
                    st.session_state.pdf_path = pdf_path_in_temp
                    st.session_state.output_filename = output_filename
                    st.session_state.api_key = current_api_key
                    st.session_state.llm_model = current_llm_model
                    st.session_state.embedding_model = current_embedding_model
                    st.session_state.selected_theme = current_selected_theme
                    st.session_state.use_persistent_index = current_use_persistent_index
                    st.session_state.force_rebuild = current_force_rebuild

                    # Change state and rerun
                    st.session_state.app_state = 'processing'
                    st.rerun()

                except FileNotFoundError as e:
                    log.error("Failed during initial file save/prep: %s", e, exc_info=True)
                    st.error(f"Error preparing for generation: {e}")
                    if temp_dir and Path(temp_dir).exists():
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    st.session_state.app_state = 'initial'
                except IOError as e:
                    log.error("Failed during initial file save/prep: %s", e, exc_info=True)
                    st.error(f"Error preparing for generation: {e}")
                    if temp_dir and Path(temp_dir).exists():
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    st.session_state.app_state = 'initial'
                except Exception as e: # Catching general exception here is okay for setup phase
                    log.error("Failed during initial file save/prep.", exc_info=True)
                    st.error(f"Error preparing for generation: {e}")
                    if temp_dir and Path(temp_dir).exists():
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    st.session_state.app_state = 'initial'


    elif st.session_state.app_state == 'processing':
        st.header("Generating Presentation")
        st.write("Please wait while the presentation is generated. This may take several minutes...")
        
        # Add progress bar
        progress_bar = st.progress(0, text="Starting generation process...")

        # --- Execute the Main Processing Logic ---
        temp_dir = st.session_state.temp_dir
        pdf_path_in_temp = st.session_state.pdf_path
        output_filename = st.session_state.output_filename
        api_key = st.session_state.api_key
        llm_model = st.session_state.llm_model
        embedding_model = st.session_state.embedding_model
        selected_theme = st.session_state.selected_theme
        use_persistent_index = st.session_state.use_persistent_index
        force_rebuild = st.session_state.force_rebuild

        try:
            # --- 1. Calculate PDF hash ---
            progress_bar.progress(10, text="Calculating PDF hash...")
            pdf_file_hash = utils.calculate_file_hash(pdf_path_in_temp)
            if not pdf_file_hash:
                log.warning("Could not calculate hash for %s. Index cache may not invalidate correctly if PDF content changes.", pdf_path_in_temp)

            # --- 2. Determine Index Storage Path ---
            progress_bar.progress(15, text="Setting up index storage...")
            index_storage_path_str = None
            if st.session_state.use_persistent_index:
                index_storage_path_str = config.PERSISTENT_INDEX_DIR
            else:
                index_storage_path_str = str(Path(temp_dir) / "index_storage")
                log.info("Using temporary index storage path: %s", index_storage_path_str)

            # --- 3. Load Documents ---
            progress_bar.progress(20, text="Loading PDF document...")
            documents = None
            with st.status("Loading PDF document...", expanded=True) as status:
                try:
                    reader = SimpleDirectoryReader(input_files=[pdf_path_in_temp])
                    documents = reader.load_data()
                    if not documents:
                        log.error("No documents loaded from %s.", pdf_path_in_temp)
                        status.update(label="Failed to load documents.", state="error")
                        raise ValueError(f"No documents loaded from {pdf_path_in_temp}. Check the file path and content.")
                    log.info("Successfully loaded %d documents.", len(documents))
                    status.update(label=f"Loaded {len(documents)} documents.", state="complete")
                except Exception as e:
                    log.error("Error loading PDF file: %s", e)
                    status.update(label="Failed to load PDF.", state="error")
                    raise e

            # --- 4. Create or Load Index ---
            progress_bar.progress(35, text="Creating or loading index...")
            index = indexing.load_or_create_index(
                storage_path_str=index_storage_path_str,
                documents=documents,
                pdf_hash=pdf_file_hash,
                api_key=api_key,
                llm_model_name=llm_model,
                embed_model_name=embedding_model,
                force_rebuild=force_rebuild
            )
            Settings.llm = GoogleGenAI(model_name=llm_model, api_key=api_key)
            Settings.embed_model = GoogleGenAIEmbedding(model_name=embedding_model, api_key=api_key)
            query_engine = index.as_query_engine(response_mode="tree_summarize")
            log.info("Query engine created.")

            # --- 5. Generate Presentation Content ---
            progress_bar.progress(50, text="Generating presentation outline with LLM...")
            with st.status("Generating presentation outline with LLM...", expanded=True) as status:
                refined_query = f"{config.FEW_SHOT_PROMPT}{config.QUERY_INSTRUCTION}"

                response = query_engine.query(refined_query)
                response_text = str(response).strip()
                log.info("Raw response received from LLM (first 500 chars):\n%s...", response_text[:500])
                status.update(label="LLM query complete.", state="complete")


            # --- 6. Clean and Parse the JSON response ---
            progress_bar.progress(70, text="Parsing LLM response...")
            json_string = None
            try:
                with st.status("Parsing LLM response...", expanded=True) as status:
                    json_string = utils.extract_json_from_response(response_text)

                    if not json_string:
                        log.error("Could not extract a potential JSON string from the LLM response.")
                        status.update(label="Failed to extract JSON.", state="error")
                        st.session_state.problematic_llm_response = response_text
                        raise ValueError("Could not extract valid JSON from LLM response.")

                    # Validate JSON structure
                    presentation_data = json.loads(json_string)
                    # Save the JSON in the temp file for debugging purpose
                    with open("debug_json.json", "w", encoding="utf-8") as debug_file:
                        debug_file.write(json_string)
                    if not isinstance(presentation_data, dict) or 'slides' not in presentation_data or not isinstance(presentation_data['slides'], list):
                        raise json.JSONDecodeError("JSON structure is invalid: missing 'slides' list or not a dictionary.", json_string, 0)

                    log.info("JSON parsed successfully.")
                    status.update(label="JSON parsed successfully.", state="complete")

            except json.JSONDecodeError as e:
                log.error("Error decoding JSON: %s", e)
                log.error("Attempted to parse:\n%s", json_string)
                status.update(label="JSON decoding failed.", state="error")
                st.session_state.problematic_json_string = json_string
                st.session_state.problematic_llm_response = response_text
                raise e
            except ValueError as e:
                log.error("Error extracting JSON: %s", e)
                status.update(label="Failed to extract JSON.", state="error")
                st.session_state.problematic_llm_response = response_text
                raise e
            except Exception as e: # Catching general exception here is okay for parsing phase
                log.error("An unexpected error occurred during JSON parsing: %s", e, exc_info=True)
                status.update(label="JSON parsing failed with unexpected error.", state="error")
                st.session_state.problematic_llm_response = response_text
                raise e


            # --- 7. Determine Config and Template Paths ---
            progress_bar.progress(85, text="Loading configuration and templates...")
            # Load the main config dictionary
            config_dict_to_use, _, config_file_exists_for_ppt, parse_successful_for_ppt = utils.load_config_file(config.CONFIG_FILE_NAME)
            
            # Update current_theme in the loaded config_dict_to_use based on the selected_theme from the UI
            if isinstance(config_dict_to_use, dict) and selected_theme:
                config_dict_to_use['current_theme'] = selected_theme
                log.info(f"Updated 'current_theme' in config_dict_to_use to '{selected_theme}' for this generation run.")
            elif not isinstance(config_dict_to_use, dict) and selected_theme:
                log.warning(f"Config data ('config_dict_to_use') is not a dictionary. Cannot set 'current_theme' to '{selected_theme}'. Default behavior or no theming might apply.")
            
            template_path_to_use = None

            # --- Logic to get template path based on the SELECTED theme from the loaded config_dict_to_use ---
            if config_file_exists_for_ppt and parse_successful_for_ppt and isinstance(config_dict_to_use, dict) and selected_theme:
                templates_dict = config_dict_to_use.get('templates')

                if isinstance(templates_dict, dict) and selected_theme in templates_dict:
                    theme_config = templates_dict[selected_theme]
                    if isinstance(theme_config, dict):
                        template_path_from_config = theme_config.get('template_path')
                        if isinstance(template_path_from_config, str) and template_path_from_config.strip():
                            template_path_str = template_path_from_config.strip()
                            # Resolve path relative to the config file's directory
                            config_dir = Path(config.CONFIG_FILE_NAME).parent
                            resolved_template_path = Path(template_path_str)
                            if not resolved_template_path.is_absolute():
                                resolved_template_path = config_dir / resolved_template_path

                            if resolved_template_path.exists():
                                template_path_to_use = str(resolved_template_path)
                                log.info("Using template path for selected theme '%s': %s", selected_theme, template_path_to_use)
                            else:
                                log.warning("Template path specified for selected theme '%s' (%s) not found at '%s'. Generating presentation without template.", selected_theme, template_path_str, resolved_template_path)
                        else:
                            log.info("Selected theme '%s' has no 'template_path' provided or path is empty. Generating presentation without template.", selected_theme)
                    else:
                        log.warning("Config structure error: Entry for selected theme '%s' in 'templates' is not a dictionary. Generating without template.", selected_theme)
                else:
                     log.warning("Selected theme '%s' not found in the 'templates' section of the config file. Generating without template.", selected_theme)
            elif selected_theme:
                 log.warning("Config file '%s' not found, invalid, or selected theme '%s' is missing. Generating without template.", config.CONFIG_FILE_NAME, selected_theme)
            else:
                 log.info("No theme selected or available. Generating presentation without template.")
            # --- End of updated logic ---


            # --- 8. Create PowerPoint Presentation ---
            progress_bar.progress(90, text=f"Creating {output_filename}...")
            output_file_path = str(Path(temp_dir) / output_filename)
            with st.status(f"Creating {output_filename}...", expanded=True) as status:
                presentation.create_presentation_from_json(
                    json_string=json_string,
                    output_path_str=output_file_path,
                    config_dict=config_dict_to_use,
                    template_path_str=template_path_to_use
                )
                status.update(label=f"{output_filename} created successfully.", state="complete")

            # --- Processing Successful ---
            progress_bar.progress(100, text="Generation complete!")
            st.session_state.output_file_path = output_file_path
            st.session_state.process_success = True
            st.session_state.success_message = f"{output_filename} generated successfully!"

        except Exception as e:
            # --- Processing Failed ---
            progress_bar.empty()
            log.error("Processing failed.", exc_info=True)
            st.session_state.process_success = False
            st.session_state.error_message = f"An error occurred during generation: {e}"
            if 'problematic_json_string' in st.session_state:
                st.session_state.error_message += "\n\nError during JSON parsing. Check the problematic JSON below."
            elif 'problematic_llm_response' in st.session_state:
                st.session_state.error_message += "\n\nError processing LLM response. Check raw response below."


        finally:
            # Transition to completed state
            st.session_state.app_state = 'completed'
            st.rerun()


    elif st.session_state.app_state == 'completed':
        st.header("Generation Complete")
        st.markdown("---")

        if st.session_state.get('process_success', False):
            st.success(st.session_state.get('success_message', "Process completed successfully."))
            output_file_path = st.session_state.get('output_file_path')
            output_filename = st.session_state.get('output_filename', 'presentation.pptx')
            pdf_path = st.session_state.get('pdf_path')

            col1, col2 = st.columns(2)
            
            with col1:
                if output_file_path and Path(output_file_path).exists():
                    try:
                        with open(output_file_path, "rb") as file_handle:
                            st.download_button(
                                label=f"Download {output_filename}",
                                data=file_handle,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True
                            )
                    except IOError as e:
                        st.error(f"Error preparing file for download: {e}")
                        log.error("Error preparing file for download: %s", e, exc_info=True)
                    except Exception as e:
                        st.error(f"Error preparing file for download: {e}")
                        log.error("Error preparing file for download: %s", e, exc_info=True)
                else:
                    st.error("Generated file not found after successful process.")
                    log.error("Generated file path '%s' from state does not exist.", output_file_path)

            with col2:
                if pdf_path and Path(pdf_path).exists():
                    try:
                        with open(pdf_path, "rb") as pdf_file:
                            pdf_bytes = pdf_file.read()
                            if st.button("View Input PDF", use_container_width=True):
                                st.subheader("Input PDF")
                                # Get temp_dir from session state
                                view_temp_dir = st.session_state.temp_dir
                                if view_temp_dir is None or not Path(view_temp_dir).exists():
                                    view_temp_dir = tempfile.mkdtemp()
                                    
                                # Create a temporary file to serve the PDF
                                pdf_display_path = Path(view_temp_dir) / "display_pdf.pdf"
                                with open(pdf_display_path, "wb") as f:
                                    f.write(pdf_bytes)
                                
                                # Display PDF using HTML iframe
                                pdf_display = f"""
                                    <iframe
                                        src="data:application/pdf;base64,{base64.b64encode(pdf_bytes).decode('utf-8')}"
                                        width="100%"
                                        height="600"
                                        type="application/pdf">
                                    </iframe>
                                """
                                st.markdown(pdf_display, unsafe_allow_html=True)
                                st.caption("PDF viewer requires browser support for embedded PDFs.")
                    except Exception as e:
                        st.error(f"Error loading input PDF: {e}")
                        log.error("Error loading input PDF: %s", e, exc_info=True)

        else:
            st.error(st.session_state.get('error_message', "Process failed."))
            if 'problematic_json_string' in st.session_state:
                st.text_area("Problematic JSON string:", st.session_state.problematic_json_string, height=300, help="This is the string that failed YAML/JSON validation.")
            if 'problematic_llm_response' in st.session_state:
                st.text_area("Raw LLM Response:", st.session_state.problematic_llm_response, height=300, help="This is the raw text received from the LLM before extraction/parsing.")

        st.markdown("---")

        if st.button("Generate Another Presentation", use_container_width=True):
            # --- Clean up temp dir ---
            if st.session_state.temp_dir and Path(st.session_state.temp_dir).exists():
                log.info("Cleaning up temp dir from completed run: %s", st.session_state.temp_dir)
                try:
                    shutil.rmtree(st.session_state.temp_dir, ignore_errors=True)
                except Exception as e: # Catching general exception here is okay for cleanup
                    log.error("Error cleaning up temp dir %s: %s", st.session_state.temp_dir, e, exc_info=True)
                    st.warning(f"Error cleaning up temporary files: {e}")
                del st.session_state.temp_dir

            # --- Clear other run-specific state variables ---
            for key in ['pdf_path', 'output_filename', 'output_file_path', 'process_success',
                        'success_message', 'error_message', 'problematic_json_string', 'problematic_llm_response',
                        'selected_theme']:
                if key in st.session_state:
                    del st.session_state[key]

            st.session_state.app_state = 'initial'
            st.session_state.page = PAGE_PRESENTATION_GENERATOR
            st.rerun()

    # Global instructions
    st.markdown("---")
    #st.write("""
   # **Instructions:**
    #1.  Enter your Google AI API Key in the sidebar.
    #2.  Upload your PDF file.
    #3.  Adjust LLM/Embedding models or output filename in the sidebar if needed.
    #4.  Use the "{0}" page to create or edit the `{1}` file. This file defines available themes and their associated template paths.
    #5.  Select the desired presentation theme from the dropdown in the sidebar. If a valid template path is associated with the selected theme in `{1}`, it will be used. Otherwise, default styling is applied.
    #6.  Check "Use Persistent Index Storage" in the sidebar to save/load the document index in the `{2}` directory. Otherwise, it's created temporarily per run.
    #7.  Click "Generate Presentation".
    #""".format(PAGE_CONFIGURATION_EDITOR, config.CONFIG_FILE_NAME, config.PERSISTENT_INDEX_DIR))


elif st.session_state.page == PAGE_CONFIGURATION_EDITOR:
    # --- Configuration Editor Page Content ---
    st.header("Edit Configuration File (`{}`)".format(config.CONFIG_FILE_NAME))
    st.write("Edit the contents of the `{}` file here. This file controls presentation styling and template options.".format(config.CONFIG_FILE_NAME))
    st.markdown("---")

    # Display messages
    if 'message' in st.session_state and st.session_state.message:
        message_type = st.session_state.pop('message_type', 'info')
        message = st.session_state.pop('message', '')
        if message:
            if message_type == "success":
                st.success(message)
            elif message_type == "info":
                st.info(message)
            elif message_type == "warning":
                st.warning(message)
            else:
                st.error(message)


    config_file_path_editor = Path(config.CONFIG_FILE_NAME)
    file_exists_now = config_file_path_editor.exists()

    # Initialize text area content
    if 'config_editor_content' not in st.session_state:
        st.session_state.config_editor_content = st.session_state.get('raw_config_text', '')


    # Check validity of current content
    current_content_in_state = st.session_state.get('config_editor_content', '')
    is_current_content_valid_yaml = False
    yaml_error_message = ""
    try:
        if not current_content_in_state and not file_exists_now:
            is_current_content_valid_yaml = True
        elif current_content_in_state:
            yaml.safe_load(current_content_in_state)
            is_current_content_valid_yaml = True
        elif not current_content_in_state and file_exists_now:
            is_current_content_valid_yaml = True

    except yaml.YAMLError as e:
        is_current_content_valid_yaml = False
        yaml_error_message = str(e)


    # --- Display status ---
    if not file_exists_now:
        st.info(f"The configuration file `{config.CONFIG_FILE_NAME}` does not exist. Enter content below or load the default, then click 'Save Configuration' to create it.")
        st.button("Load Default Content Template", on_click=load_default_config)

    elif not is_current_content_valid_yaml:
        st.warning(f"The content for `{config.CONFIG_FILE_NAME}` is currently invalid YAML. Please correct the errors below before saving.")
        if yaml_error_message:
            st.code(yaml_error_message, language='text')

    else:
        st.success(f"Editing `{config.CONFIG_FILE_NAME}`. Content is valid YAML. Save changes to apply.")


    # Text area for editing
    st.text_area(
        "Configuration Content (YAML)",
        height=600,
        key='config_editor_content'
    )

    col1, col2, col3 = st.columns([1, 1, 2])

    with col1:
        st.button(
            "Save Configuration",
            use_container_width=True,
            type="primary",
            on_click=save_config_to_file,
            disabled=not is_current_content_valid_yaml
        )

    with col2:
        st.button(
            "Delete Configuration File",
            use_container_width=True,
            type="secondary",
            on_click=delete_config_file,
            disabled=not file_exists_now
        )

    with col3:
        st.button(
            "Reload from File",
            use_container_width=False,
            on_click=reload_config_from_file,
            disabled=not file_exists_now
        )


    st.markdown("---")
    st.write("The config file allows you to override default styling and specify a custom template.")
    st.write(f"If `{config.CONFIG_FILE_NAME}` exists and is valid YAML, the settings within it will be applied when generating presentations.")
    st.write("If it does not exist or contains invalid YAML, default styling will be used, and the 'Use Template File' option will be ignored.")
    st.write("Refer to the comments within the default content (available via 'Load Default Content Template' above) for guidance on keys and structure.")


