import logging
from pathlib import Path
from template_pptgen import create_ppt_with_template
from default_pptgen import create_ppt_without_template

log = logging.getLogger(__name__)

# --- Presentation Creation Function ---

def create_presentation_from_json(json_string: str, output_path_str: str, config_dict: dict | None, template_path_str: str | None):
    """Create a PowerPoint presentation from a JSON string, using config dict and template path.

    Args:
        json_string (str): The JSON string containing the slide data.
        output_path_str (str): The desired path for the output .pptx file.
        config_dict (dict | None): Dictionary containing configuration settings. If None, default config will be loaded.
        template_path_str (str | None): Path to the .pptx template file (optional).
    """
    output_path = Path(output_path_str)

    # Resolve template path relative to script dir if not absolute and exists
    template_path = None
    if template_path_str:
        p = Path(template_path_str)
        # Templates might be specified relative to config or script, assume script for now if relative
        if not p.is_absolute():
            p = Path(__file__).parent / p
        if p.exists():
            template_path = p
        else:
            log.warning("Provided template path '%s' (resolved to '%s') does not exist. Ignoring.", template_path_str, p)

    # Use the provided config dictionary or load default if None
    final_config_dict = config_dict if config_dict is not None else None

    log.info("Attempting to create presentation: %s", output_path)
    log.info("Using provided config dictionary: %s", "Yes" if config_dict else "No (using defaults)")
    log.info("Using template path: %s (Exists: %s)", template_path, template_path.exists() if template_path else False)

    try:
        # Ensure the output directory exists
        output_dir = output_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        log.info("Ensured output directory exists: %s", output_dir)

        template_path_to_pass = str(template_path) if template_path else None

        log.info("Calling presentation creation with template: %s", template_path_to_pass)

        if template_path_to_pass:
            log.info("Valid template found. Creating PowerPoint with template.")
            # Ensure the wrapper function exists and is callable
            if callable(create_ppt_with_template):
                create_ppt_with_template(
                    json_string,
                    output_path=str(output_path),
                    config_dict=final_config_dict,
                    template_path=template_path_to_pass
                )
            else:
                raise TypeError("create_ppt_with_template is not callable")
        else:
            log.info("No valid template path provided or found. Creating PowerPoint without template.")
            # Ensure the function exists and is callable
            if callable(create_ppt_without_template):
                create_ppt_without_template(
                   json_input=json_string,
                   output_ppt_file=str(output_path),
                   config_dict=final_config_dict
                )
            else:
                raise TypeError("create_ppt_without_template is not callable")

        log.info("Presentation file created successfully at %s.", output_path)

    except FileNotFoundError as e:
        log.error("File not found during presentation creation: %s. Ensure template (%s) file exists if needed.", e, template_path)
        # Error should be handled by the caller (Streamlit app)
        raise e
    except Exception as e:
        log.error("An error occurred during presentation creation: %s", e, exc_info=True)
        # Error should be handled by the caller (Streamlit app)
        raise e
