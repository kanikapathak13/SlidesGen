import random
import sys
import requests
from PIL import Image
import io
from duckduckgo_search import DDGS
import os
import hashlib 
from typing import Optional, Tuple, List
import json
from urllib.parse import urlencode
from dotenv import load_dotenv
import nltk
from nltk.corpus import wordnet
import re

# Load environment variables
load_dotenv()

# Download required NLTK data
try:
    nltk.data.find('corpora/wordnet')
except LookupError:
    nltk.download('wordnet')

def get_synonyms(word: str) -> List[str]:
    """Get synonyms for a word from WordNet."""
    synonyms = set()
    synsets = wordnet.synsets(word)
    for synset in synsets[:2]:  # Use first two synsets
        synonyms.update(lemma.name() for lemma in synset.lemmas())
    return [s for s in synonyms if s != word and '_' not in s]

def get_query_variations(query: str) -> List[str]:
    """Get common variations of the query by adding descriptive terms."""
    variations = []
    
    # Add image-related terms if not already present
    if not any(word in query.lower() for word in ['picture', 'image', 'photo']):
        variations.extend([f"{query} picture", f"{query} image", f"{query} photo"])
    
    # Add diagram variations
    if 'diagram' in query.lower():
        variations.extend([f"{query} illustration", f"{query} schematic", f"{query} visual"])
    
    # Add chart/graph variations
    elif any(word in query.lower() for word in ['chart', 'graph']):
        variations.extend([f"{query} visualization", f"{query} data visualization", f"{query} infographic"])
    
    return variations

def get_shorter_queries(query: str) -> List[str]:
    """Generate shorter versions of the query by removing less important words."""
    words = query.split()
    
    # If query is already short, don't shorten further
    if len(words) <= 2:
        return []
    
    shorter_queries = []
    
    # Remove articles, conjunctions, and prepositions
    stopwords = ['a', 'an', 'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'with', 'by', 'about', 'as']
    filtered_words = [word for word in words if word not in stopwords]
    
    # If we have a meaningful filtered query, add it
    if len(filtered_words) >= 2 and len(filtered_words) < len(words):
        shorter_queries.append(' '.join(filtered_words))
    
    # Extract key phrases: first half, second half, middle segment
    if len(words) >= 4:
        midpoint = len(words) // 2
        shorter_queries.append(' '.join(words[:midpoint]))
        shorter_queries.append(' '.join(words[midpoint:]))
        
        # If query is long enough, extract middle segment
        if len(words) >= 6:
            start = max(0, midpoint - 1)
            end = min(len(words), midpoint + 2)
            shorter_queries.append(' '.join(words[start:end]))
    
    # Take the first few essential words if query is long
    if len(words) > 3:
        shorter_queries.append(' '.join(words[:3]))
    
    # Remove duplicates and the original query
    shorter_queries = list(set(shorter_queries))
    if query in shorter_queries:
        shorter_queries.remove(query)
    
    return shorter_queries

def generate_alternative_queries(query: str) -> List[str]:
    """
    Generate alternative search queries using synonyms and related terms.
    
    Args:
        query (str): Original search query
        
    Returns:
        List[str]: List of alternative search queries
    """
    # Clean the query
    query = query.strip().lower()
    words = query.split()
    
    alternatives = []
    
    # Generate word-replacement alternatives
    for word in words:
        synonyms = get_synonyms(word)
        for synonym in synonyms:
            new_query = ' '.join(w if w != word else synonym for w in words)
            alternatives.append(new_query)
    
    # Add common variations based on query type
    alternatives.extend(get_query_variations(query))
    
    # Add shorter fallback queries
    alternatives.extend(get_shorter_queries(query))
    
    # Remove duplicates and the original query
    alternatives = list(set(alternatives))
    if query in alternatives:
        alternatives.remove(query)
    
    # Limit the number of alternatives
    return alternatives[:5]  # Return top 5 alternatives

def get_image(query: str, save_dir: str = 'images', max_attempts: int = 6):
    """
    Searches for an image using multiple sources and queries, attempts to download and process one
    from a list of candidates, saves it to disk, and returns the PIL Image object, its save path, 
    aspect ratio, and the URL of the downloaded image.

    Args:
        query (str): The search query for the image.
        save_dir (str): Directory to save the downloaded image.
        max_attempts (int): Maximum number of image URLs to try from each source.

    Returns:
        tuple: (PIL.Image.Image, str, float, str) containing 
               the PIL image object,
               the path to the saved image (or None if save failed), 
               the aspect ratio (height/width),
               and the URL of the successfully downloaded image (or None if all attempts fail).
               Returns (None, None, None, None) if all attempts fail.
    """
    os.makedirs(save_dir, exist_ok=True)

    # Start with the original query
    queries_to_try = [query]
    
    # Generate alternative queries
    alternative_queries = generate_alternative_queries(query)
    queries_to_try.extend(alternative_queries)
    
    for current_query in queries_to_try:
        print(f"Trying query: '{current_query}'", file=sys.stderr)
        
        # Try each image source in sequence
        all_candidate_urls = []
        
        # 1. Try DuckDuckGo
        ddg_urls = get_image_from_duckduckgo(current_query)
        all_candidate_urls.extend(ddg_urls)
        
        # 2. If DuckDuckGo failed or returned no results, try Unsplash
        if not ddg_urls:
            print("DuckDuckGo returned no results, trying Unsplash...", file=sys.stderr)
            unsplash_urls = get_image_from_unsplash(current_query)
            all_candidate_urls.extend(unsplash_urls)
        
        # 3. If both failed or returned no results, try Google Custom Search
        if not all_candidate_urls:
            print("Unsplash returned no results, trying Google Custom Search...", file=sys.stderr)
            google_urls = get_image_from_google(current_query)
            all_candidate_urls.extend(google_urls)

        if all_candidate_urls:
            # Shuffle all collected URLs
            random.shuffle(all_candidate_urls)
            
            # Try processing images from the collected URLs
            for attempt_num, current_url in enumerate(all_candidate_urls[:max_attempts]):
                print(f"Attempt {attempt_num + 1}/{min(len(all_candidate_urls), max_attempts)} for query '{current_query}'. Trying URL: {current_url}", file=sys.stderr)
                
                result = process_image_from_url(current_url, current_query, save_dir)
                if result[0] is not None:  # If processing succeeded
                    return result
        
        print(f"No suitable images found for query: '{current_query}'", file=sys.stderr)
    
    print(f"All queries and attempts failed for original query: '{query}'", file=sys.stderr)
    return None, None, None, None

def get_image_from_duckduckgo(query: str, max_results: int = 15) -> List[str]:
    """Get image URLs from DuckDuckGo search."""
    try:
        with DDGS() as ddgs:
            # Add type_image='photo' to prefer actual photos over clipart/drawings
            # Add size='Large' to get better quality images
            search_results = list(ddgs.images(
                query,
                max_results=max_results,
                type_image='photo',
                size='Large'
            ))
        
        candidate_urls = []
        for result_item in search_results:
            url = result_item.get('image')
            if url:
                candidate_urls.append(url)
        return candidate_urls
    except Exception as e:
        print(f"DuckDuckGo search failed: {e}", file=sys.stderr)
        return []

def get_image_from_unsplash(query: str, max_results: int = 15) -> List[str]:
    """Get image URLs from Unsplash API."""
    unsplash_api_key = os.getenv('UNSPLASH_API_KEY')
    if not unsplash_api_key:
        print("Unsplash API key not found in environment variables.", file=sys.stderr)
        return []
    
    try:
        headers = {'Authorization': f'Client-ID {unsplash_api_key}'}
        # Add orientation=landscape to get better aspect ratios for slides
        params = {
            'query': query,
            'per_page': max_results,
            'orientation': 'landscape'
        }
        response = requests.get(
            'https://api.unsplash.com/search/photos',
            headers=headers,
            params=params
        )
        response.raise_for_status()
        results = response.json()
        return [photo['urls']['regular'] for photo in results.get('results', [])]
    except Exception as e:
        print(f"Unsplash search failed: {e}", file=sys.stderr)
        return []

def get_image_from_google(query: str, max_results: int = 15) -> List[str]:
    """Get image URLs from Google Custom Search API."""
    google_api_key = os.getenv('GOOGLE_API_KEY')
    google_cx = os.getenv('GOOGLE_CX')
    
    if not google_api_key or not google_cx:
        print("Google API key or Custom Search Engine ID not found in environment variables.", file=sys.stderr)
        return []
    
    try:
        base_url = "https://www.googleapis.com/customsearch/v1"
        params = {
            'key': google_api_key,
            'cx': google_cx,
            'q': query,
            'searchType': 'image',
            'num': min(max_results, 10),  # Google CSE has a max of 10 results per query
            'imgType': 'photo',  # Prefer photos over other types
            'imgSize': 'large',  # Get high quality images
            'safe': 'active'  # Enable safe search
        }
        
        response = requests.get(base_url, params=params)
        response.raise_for_status()
        results = response.json()
        return [item['link'] for item in results.get('items', [])]
    except Exception as e:
        print(f"Google Custom Search failed: {e}", file=sys.stderr)
        return []

def process_image_from_url(url: str, query: str, save_dir: str) -> Tuple[Optional[Image.Image], Optional[str], Optional[float], Optional[str]]:
    """Process a single image URL and return the image object, save path, aspect ratio, and URL."""
    try:
        response = requests.get(url, stream=True, timeout=15,
                              headers={'User-Agent': 'Mozilla/5.0'})
        response.raise_for_status()
        
        image_content = response.content
        pil_image_object = None
        image_format_str = 'JPEG'
        
        with io.BytesIO(image_content) as image_bytes_io:
            with Image.open(image_bytes_io) as img:
                img.verify()
                image_bytes_io.seek(0)
                pil_image_object = Image.open(image_bytes_io)
                pil_image_object.load()

        if not pil_image_object:
            return None, None, None, None

        # Check minimum dimensions for better quality
        width, height = pil_image_object.size
        if width < 800 or height < 600:  # Skip images that are too small
            return None, None, None, None

        image_format_str = pil_image_object.format or 'JPEG'
        
        if image_format_str.upper() == 'WEBP':
            if pil_image_object.mode == 'RGBA':
                pil_image_object = pil_image_object.convert('RGB')
            image_format_str = 'JPEG'

        aspect_ratio = height / width
        # Skip images with extreme aspect ratios
        if aspect_ratio < 0.5 or aspect_ratio > 2.0:
            return None, None, None, None

        safe_query_part = "".join(c if c.isalnum() or c in ['_','-'] else '_' for c in query.replace(' ', '_'))
        safe_query_part = safe_query_part[:50]
        url_hash = hashlib.sha1(url.encode('utf-8')).hexdigest()[:8]
        file_extension = image_format_str.lower()
        if not file_extension or len(file_extension) > 5 or file_extension == "webp":
            file_extension = "jpg"
        
        base_file_name = f"{safe_query_part}_{url_hash}"
        image_file_name = f"{base_file_name}.{file_extension}"
        image_save_path = os.path.join(save_dir, image_file_name)
        
        # Save the image
        pil_image_object.save(image_save_path, format=image_format_str)
        
        return pil_image_object, image_save_path, aspect_ratio, url
        
    except Exception as e:
        print(f"Failed to process image from URL {url}: {e}", file=sys.stderr)
        return None, None, None, None

# # --- Example Usage ---
# if __name__ == "__main__":
#     search_query_1 = "modern minimalist living room"
#     search_query_2 = "abstract geometric patterns"
#     search_query_3 = "futuristic cityscape" # A new query

#   #  queries_to_test = [search_query_1, search_query_2, search_query_3] 

#     for i, current_query in enumerate(queries_to_test):
#         print(f"\n--- Processing query {i+1}: '{current_query}' ---")
#         pil_obj, file_path, ar, img_url = get_image(current_query)
        
#         if pil_obj:
#             print(f"  Successfully processed image for '{current_query}'.")
#             print(f"  Downloaded from URL: {img_url}")
#             print(f"  PIL Image Object: Present (Format: {pil_obj.format}, Size: {pil_obj.size}, Mode: {pil_obj.mode})")
#             print(f"  Aspect Ratio: {ar:.2f}")
#             if file_path:
#                 print(f"  Saved to: {file_path}")
#             else:
#                 print(f"  Image processed but FAILED to save to disk.")
#         else:
#             print(f"  Failed to download or process image for query: '{current_query}'")

#     # No longer tracking used URLs globally in this example
#     print("\n--- Example script finished ---")
