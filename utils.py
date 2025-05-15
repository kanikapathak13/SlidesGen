import logging
import re
import yaml
import hashlib
from pathlib import Path
import streamlit as st # Needed for save_uploaded_file type hint

log = logging.getLogger(__name__)

# --- File Handling ---

# @st.cache_data(show_spinner=False) # Caching might be problematic if file content changes but name doesn't
def save_uploaded_file(uploaded_file: st.runtime.uploaded_file_manager.UploadedFile, target_dir: str, filename: str | None = None) -> str | None:
    """Saves an uploaded file to a target directory."""
    if not uploaded_file:
        log.warning("save_uploaded_file called with no uploaded file.")
        return None

    target_path = Path(target_dir)
    target_path.mkdir(parents=True, exist_ok=True) # Ensure directory exists

    save_filename = filename or uploaded_file.name
    full_path = target_path / save_filename

    try:
        with open(full_path, "wb") as f:
            f.write(uploaded_file.getvalue())
        log.info("File '%s' saved successfully to '%s'.", save_filename, full_path)
        return str(full_path)
    except Exception as e:
        log.error("Error saving uploaded file '%s' to '%s': %s", save_filename, full_path, e, exc_info=True)
        st.error(f"Error saving file {save_filename}: {e}")
        return None

# @st.cache_data(show_spinner=False) # Caching based on path might be okay, but hash depends on content
def calculate_file_hash(file_path: str | Path) -> str | None:
    """Calculates the SHA256 hash of a file."""
    try:
        hasher = hashlib.sha256()
        with open(file_path, 'rb') as file:
            while True:
                chunk = file.read(4096) # Read in chunks
                if not chunk:
                    break
                hasher.update(chunk)
        return hasher.hexdigest()
    except FileNotFoundError:
        log.warning("File not found for hashing: %s", file_path)
        return None
    except Exception as e:
        log.error("Error calculating hash for file '%s': %s", file_path, e, exc_info=True)
        return None

# --- Configuration Loading ---

# @st.cache_data(show_spinner=False) # Cache based on filename, might need invalidation if file changes
def load_config_file(config_filename: str) -> tuple[dict | None, str | None, bool, bool]:
    """Loads and parses the YAML configuration file.

    Returns:
        tuple[dict | None, str | None, bool, bool]:
            config_data: Parsed YAML content as dict, or None on error/not found.
            raw_content: Raw text content of the file, or None if not found/read error.
            file_exists: Boolean indicating if the file exists.
            parse_successful: Boolean indicating if YAML parsing succeeded.
    """
    config_path = Path(config_filename)
    raw_content = None
    config_data = None
    parse_successful = False

    if not config_path.exists():
        log.info("Config file '%s' not found.", config_filename)
        return None, None, False, False

    try:
        raw_content = config_path.read_text(encoding='utf-8')
        log.info("Read config file '%s'.", config_filename)
        try:
            config_data = yaml.safe_load(raw_content)
            if isinstance(config_data, dict):
                log.info("Config file '%s' parsed successfully.", config_filename)
                parse_successful = True
            else:
                log.warning("Config file '%s' content is not a valid dictionary.", config_filename)
                config_data = None
                parse_successful = False
        except yaml.YAMLError as e:
            log.error("Error parsing YAML in config file '%s': %s", config_filename, e, exc_info=True)
            config_data = None
            parse_successful = False
        except Exception as e:
             log.error("Unexpected error parsing YAML in config file '%s': %s", config_filename, e, exc_info=True)
             config_data = None
             parse_successful = False

    except IOError as e:
        log.error("Error reading config file '%s': %s", config_filename, e, exc_info=True)
        raw_content = None
        config_data = None
        parse_successful = False
    except Exception as e:
        log.error("Unexpected error loading config file '%s': %s", config_filename, e, exc_info=True)
        raw_content = None
        config_data = None
        parse_successful = False

    return config_data, raw_content, True, parse_successful


# --- JSON Extraction ---

def extract_json_from_response(response_text: str) -> str | None:
    """Attempts to extract a JSON object or array string from LLM response text.

    It first looks for JSON within markdown code blocks (```json ... ```).
    If not found, it looks for the outermost curly braces {} or square brackets [].
    """
    log.debug("Attempting to extract JSON from response (first 100 chars): %s...", response_text[:100])

    # 1. Check for JSON within markdown code blocks
    json_match = re.search(r"```json\s*([\s\S]*?)\s*```", response_text, re.DOTALL)
    if json_match:
        potential_json = json_match.group(1).strip()
        log.info("Found potential JSON inside markdown code block.")
        if (potential_json.startswith('{') and potential_json.endswith('}')) or \
           (potential_json.startswith('[') and potential_json.endswith(']')):
            return potential_json
        else:
            log.warning("Content in markdown block doesn't start/end like JSON. Searching outside block.")

    # 2. If no valid markdown block found, search for the first '{' or '['
    #    and the last '}' or ']' in the entire response.
    first_bracket = response_text.find('[')
    first_brace = response_text.find('{')

    start_index = -1
    if first_brace != -1 and first_bracket != -1:
        start_index = min(first_brace, first_bracket)
    elif first_brace != -1:
        start_index = first_brace
    elif first_bracket != -1:
        start_index = first_bracket

    if start_index == -1:
        log.warning("No starting '{' or '[' found in the response.")
        return None

    last_brace = response_text.rfind('}')
    last_bracket = response_text.rfind(']')
    end_index = -1

    possible_ends = []
    if last_brace > start_index:
        possible_ends.append(last_brace)
    if last_bracket > start_index:
        possible_ends.append(last_bracket)

    if not possible_ends:
        log.warning("Found a starting bracket/brace but no corresponding ending one after it.")
        return None

    end_index = max(possible_ends) + 1

    potential_json = response_text[start_index:end_index].strip()
    log.info("Extracted potential JSON substring based on first/last brackets/braces.")
    if (potential_json.startswith('{') and potential_json.endswith('}')) or \
       (potential_json.startswith('[') and potential_json.endswith(']')):
        return potential_json
    else:
        log.warning("Extracted substring doesn't start/end like JSON, despite finding brackets/braces.")
        return None

    log.warning("Could not extract a JSON-like string from the response.")
    return None
