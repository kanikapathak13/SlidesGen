import streamlit as st
import logging
import shutil
from pathlib import Path

from llama_index.core import Settings, StorageContext, load_index_from_storage, VectorStoreIndex, Document
from llama_index.llms.google_genai import GoogleGenAI
from llama_index.embeddings.google_genai import GoogleGenAIEmbedding

import config
log = logging.getLogger(__name__)


# Use Streamlit's caching. Include pdf_hash and model names in the cache key
# so the cache updates if the PDF or models change.
# Documents are hashed using llama_index's built-in hash method.
@st.cache_resource(show_spinner=False, hash_funcs={Document: lambda d: d.hash})
def load_or_create_index(
    storage_path_str: str,
    documents: list[Document],
    pdf_hash: str | None,
    api_key: str,
    llm_model_name: str,
    embed_model_name: str,
    force_rebuild: bool = False
    ) -> VectorStoreIndex:
    """Loads an index from storage if it exists and matches the current PDF hash.
    Otherwise, creates a new index and saves it.

    Uses Streamlit's caching to speed things up. Checks the PDF hash in
    persistent storage to decide if the index needs rebuilding.
    """
    log.info("Starting load_or_create_index for path: %s (Force rebuild: %s)", storage_path_str, force_rebuild)

    storage_path = Path(storage_path_str)
    is_persistent = (storage_path_str == config.PERSISTENT_INDEX_DIR)

    # --- Configure LlamaIndex Settings ---
    # Set models globally for this execution context.
    try:
        Settings.llm = GoogleGenAI(model_name=llm_model_name, api_key=api_key)
        Settings.embed_model = GoogleGenAIEmbedding(model_name=embed_model_name, api_key=api_key)
        log.info("LlamaIndex Settings configured: LLM='%s', Embed='%s'", llm_model_name, embed_model_name)
    except Exception as e:
        log.error("Failed to initialize Google AI models: %s", e, exc_info=True)
        st.error(f"Error initializing Google AI models: {e}. Please check API key and model names.")
        raise

    # --- Check for Existing Index and Hash Match (if persistent) ---
    index_exists = False
    hash_match = True # Assume match unless persistent storage proves otherwise
    hash_file = storage_path / "source_doc.hash"

    if not force_rebuild and storage_path.exists() and any(storage_path.iterdir()):
        index_exists = True
        log.info("Index storage path '%s' exists and is not empty.", storage_path)
        if is_persistent and pdf_hash:
            if hash_file.exists():
                try:
                    stored_hash = hash_file.read_text().strip()
                    if stored_hash != pdf_hash:
                        hash_match = False
                        log.warning("PDF hash mismatch! Stored: '%s', Current: '%s'. Index will be rebuilt.", stored_hash, pdf_hash)
                    else:
                        log.info("PDF hash matches stored hash.")
                except IOError as e:
                    log.warning("Could not read or compare hash file '%s': %s. Assuming mismatch.", hash_file, e, exc_info=True)
                    hash_match = False # Rebuild if hash file is unreadable or other error
                except Exception as e:
                    log.warning("Error comparing hash for file '%s': %s. Assuming mismatch.", hash_file, e, exc_info=True)
                    hash_match = False
            else:
                log.warning("Hash file '%s' not found in persistent storage. Assuming mismatch.", hash_file)
                hash_match = False # Rebuild if hash file is missing
        elif is_persistent and not pdf_hash:
             log.warning("No PDF hash provided for persistent index check. Loading existing index without verification.")
             # Proceed to load, but it might be stale.

    # --- Decide whether to Load or Create ---
    index = None
    if index_exists and hash_match and not force_rebuild:
        with st.status(f"Loading existing index from {storage_path_str}...", expanded=True) as status:
            log.info("Attempting to load index from %s...", storage_path)
            try:
                storage_context = StorageContext.from_defaults(persist_dir=str(storage_path))
                index = load_index_from_storage(storage_context)
                log.info("Index loaded successfully.")
                status.update(label="Index loaded successfully.", state="complete")
            except Exception as e:
                log.error("Error loading index from storage: %s. Will attempt to rebuild.", e, exc_info=True)
                status.update(label=f"Error loading index: {e}. Rebuilding...", state="error")
                index = None # Signal that rebuild is needed
                # Clean up potentially corrupted storage before rebuilding
                if is_persistent:
                    log.warning("Removing potentially corrupted persistent index directory: %s", storage_path)
                    shutil.rmtree(storage_path, ignore_errors=True)

    if index is None: # Rebuild if loading failed, forced, doesn't exist, or hash mismatch
        action = "Creating new index"
        if force_rebuild: action = "Forcibly rebuilding index"
        elif not index_exists: action = "Creating new index (storage path not found or empty)"
        elif not hash_match: action = "Rebuilding index (PDF hash mismatch)"
        # The case where loading failed is implicitly covered as index is None

        with st.status(f"{action} in {storage_path_str}...", expanded=True) as status:
            log.info("%s...", action)
            try:
                # Ensure storage path exists
                storage_path.mkdir(parents=True, exist_ok=True)

                # Create index from documents
                index = VectorStoreIndex.from_documents(
                    documents,
                    show_progress=True,
                    # Uses Settings.llm and Settings.embed_model configured earlier
                )
                status.update(label="Index created.", state="running")

                # Persist the index
                log.info("Persisting index to %s...", storage_path)
                index.storage_context.persist(persist_dir=str(storage_path))
                status.update(label="Index persisted.", state="running")

                # Store the hash if persistent and hash is available
                if is_persistent and pdf_hash:
                    try:
                        hash_file.write_text(pdf_hash)
                        log.info("Stored current PDF hash '%s' to '%s'.", pdf_hash, hash_file)
                    except IOError as e:
                        log.warning("Failed to write hash file '%s': %s", hash_file, e, exc_info=True)
                        st.warning("Index created, but failed to store document hash. Future checks might trigger unnecessary rebuilds.")
                    except Exception as e:
                        log.warning("Unexpected error writing hash file '%s': %s", hash_file, e, exc_info=True)
                        st.warning("Index created, but failed to store document hash. Future checks might trigger unnecessary rebuilds.")


                log.info("Index creation and persistence complete.")
                status.update(label=f"{action} complete.", state="complete")

            except Exception as e:
                log.error("Error during index creation or persistence: %s", e, exc_info=True)
                status.update(label=f"Error during index creation: {e}", state="error")
                # Clean up potentially incomplete index directory
                if storage_path.exists():
                    log.warning("Cleaning up potentially incomplete index directory: %s", storage_path)
                    try:
                        shutil.rmtree(storage_path, ignore_errors=True)
                    except Exception as cleanup_e:
                        log.error("Error during cleanup of index directory: %s", cleanup_e)
                st.error(f"Failed to create or persist index: {e}")
                raise

    if index is None:
         # This state indicates a critical failure in both loading and creation paths.
         log.error("Index is None after load/create process. This indicates a critical failure.")
         st.error("Failed to load or create the document index. Cannot proceed.")
         raise RuntimeError("Index could not be loaded or created.")

    return index
