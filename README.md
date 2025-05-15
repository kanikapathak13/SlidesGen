# Image Search with Multiple Sources

This module provides image search functionality using multiple sources:
1. DuckDuckGo (no API key required)
2. Unsplash API (requires API key)
3. Google Custom Search API (requires API key and Custom Search Engine ID)

## Setup

1. Install the required dependencies:
```bash
pip install python-dotenv requests Pillow duckduckgo_search
```

2. Create a `.env` file in your project root with the following variables:
```env
# Unsplash API credentials
# Get from https://unsplash.com/developers
UNSPLASH_API_KEY=your_unsplash_access_key_here

# Google Custom Search API credentials
# Get from https://developers.google.com/custom-search/v1/introduction
GOOGLE_API_KEY=your_google_api_key_here
# Get from https://programmablesearchengine.google.com/
GOOGLE_CX=your_google_custom_search_engine_id_here
```

## Usage

The module will try each image source in sequence:
1. First attempts to find images using DuckDuckGo
2. If DuckDuckGo fails or returns no results, tries Unsplash
3. If both fail, falls back to Google Custom Search

Example usage:

```python
from image_handler import get_image

# Search for an image
image, save_path, aspect_ratio, url = get_image(
    query="modern minimalist living room",
    save_dir="images",
    max_attempts=4
)

if image:
    print(f"Image saved to: {save_path}")
    print(f"Image aspect ratio: {aspect_ratio}")
    print(f"Source URL: {url}")
else:
    print("Failed to find/download image")
```

## API Keys

### Unsplash
1. Sign up at https://unsplash.com/developers
2. Create a new application
3. Copy the Access Key to your `.env` file

### Google Custom Search
1. Create a project in Google Cloud Console
2. Enable Custom Search API
3. Create API credentials and copy the API key
4. Go to https://programmablesearchengine.google.com/
5. Create a new search engine
6. Enable "Image Search"
7. Copy the Search Engine ID (cx) to your `.env` file 