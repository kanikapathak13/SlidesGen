import os

CONFIG_FILE_NAME = "slide_config.yaml"
PERSISTENT_INDEX_DIR = os.path.join(os.path.dirname(__file__), "index_storage") # Store index next to script

# --- Default Configuration Content ---
DEFAULT_CONFIG_CONTENT = """\
# GenSlides Configuration File (YAML Format)
# This file controls the appearance and structure of the generated PowerPoint.

# --- Template Path ---
# Optional: Specify the path to a .pptx template file.
# If relative, it's path relative to this config file's location.
# If commented out or blank, no template will be used (defaults applied).
# Example: template_path: "templates/MyCustomTemplate.pptx"
template_path: "templates/Organic presentation.pptx"

# --- Font Settings (Optional) ---
# Specify default fonts for different text types.
# If omitted, defaults from the template (if used) or PowerPoint's defaults apply.
# fonts:
#   title: "Calibri Light"
#   body: "Calibri"
#   size_title: 32  # In points
#   size_body: 18   # In points

# --- Color Settings (Optional) ---
# Define colors using RGB hex codes (e.g., "FF0000" for red).
# If omitted, defaults from the template (if used) or PowerPoint's defaults apply.
# colors:
#   background: "FFFFFF" # Slide background
#   text: "000000"       # Default text color
#   accent1: "4472C4"    # Example accent color

# --- Slide Master & Layout Defaults (Advanced) ---
# Usually controlled by the template file. Modify with caution.
# slide_master_index: 0 # Index of the slide master to use (usually 0)
# layout_mapping: # Maps JSON layout_idx to PPT layout index in the template
#   0: 0  # Title Slide
#   1: 1  # Title and Content
#   2: 2  # Section Header
#   3: 3  # Two Content
#   4: 4  # Comparison
#   5: 5  # Title Only
#   6: 6  # Blank
#   7: 7  # Content with Caption
#   8: 8  # Picture with Caption
#   # Add more if your template has custom layouts you want to map

# --- Other Options ---
# add_slide_numbers: true # Add slide numbers to the footer (requires placeholder in template)
"""

# --- LLM Prompt for Presentation Generation ---
FEW_SHOT_PROMPT = """\
You are an expert AI assistant specialized in transforming long document texts into highly structured, concise, and **visually engaging** PowerPoint presentation outlines, making strategic use of standard layouts, including those designed for visuals (layouts 7 & 8) when appropriate. Your goal is to analyze the provided document, identify key themes, sections, details, definitions, comparisons, arguments, and **opportunities for visual representation**, and organize them into a logical slide sequence optimized for clarity and visual appeal.

**CRITICAL INSTRUCTIONS FOR COVERAGE AND STRUCTURE:**

1.  **COMPREHENSIVE COVERAGE:** Your presentation outline MUST cover **all essential and important topics, concepts, methods, findings, and conclusions** discussed in the document. Do **NOT** just provide a high-level summary. Aim for thoroughness.
2.  **LOGICAL FLOW:** Structure the slides logically, ideally following the natural flow of the original document (e.g., Introduction, Background, Methodology, Results, Discussion, Conclusion). Use Section Header slides (`layout_idx: 2`) to clearly demarcate major parts.
3.  **ADEQUATE LENGTH:** For longer documents, ensure you generate a sufficient number of slides to cover the content without rushing. Do not limit the presentation arbitrarily to a few pages if the document is extensive.
4.  **BEAUTIFUL CONCLUSION:** Include a dedicated concluding slide (`layout_idx: 1` or similar) that summarizes the main takeaways, the significance of the work, and potentially mention limitations or future directions. This should provide a strong, concluding summary.

The output MUST be a single, valid JSON object. This object MUST contain a top-level key named `"slides"`, which holds a list of slide objects.

**Critical Rules for JSON Output (Adhere Strictly):**

1.  **Mandatory `layout_idx` (0-8):** Every slide object MUST have a `"layout_idx"` key with an integer value from **0 to 8**, corresponding to the standard PowerPoint layouts. Choose the MOST appropriate layout for the content.
2.  **Layout-Specific Fields ONLY:** Include ONLY the fields relevant to the chosen `"layout_idx"`. Do NOT invent fields or include fields for placeholders absent on that layout.
    *   `layout_idx: 0` (Title Slide): Requires `"title"` (str), `"subtitle"` (str).
    *   `layout_idx: 1` (Title and Content): Requires `"title"` (str), `"content"` (list of str / single str).
    *   `layout_idx: 2` (Section Header): Requires `"section_title"` (str). Optional: `"section_description"` (str).
    *   `layout_idx: 3` (Two Content): Requires `"title"` (str), `"left_content"` (list/str), `"right_content"` (list/str).
    *   `layout_idx: 4` (Comparison): Requires `"title"` (str), `"left_heading"` (str), `"right_heading"` (str), `"left_comparison_content"` (list/str), `"right_comparison_content"` (list/str).
    *   `layout_idx: 5` (Title Only): Requires only `"title"` (str).
    *   `layout_idx: 6` (Blank): Requires NO specific content fields besides `"layout_idx": 6`. Use for custom visuals (describe in notes).
    *   `layout_idx: 7` (Content with Caption): Requires `"title"` (str), `"caption_text"` (list/str), `"object_description"` (str - describe the visual).
    *   `layout_idx: 8` (Picture with Caption): Requires `"picture_description"` (str - describe the image), `"caption_text"` (str - short caption for image).
3.  **Content Field Structure:** Fields like `content`, `left_content`, etc., should be a **list of strings** for multiple bullets/paragraphs. A single string is acceptable if the placeholder is for one paragraph.
4.  **Markdown Formatting:** Use `**Bold**`, `*Italic*`, and `<u>Underline` within ALL text strings for emphasis and clarity.
5.  **Indentation:** Use exactly **two leading spaces** (`  `) for each indentation level within content lists (e.g., `"  - Sub-point"`).
6.  **Appropriate Layout Choice:** *Strategically* choose the best layout for the content. While layouts like Title and Content (1), Two Content (3), and Comparison (4) are crucial for organizing text and lists, **make an effort to identify content that would be significantly enhanced by a visual aid.** When the document describes a concept, diagram, process, data visualization (charts/graphs), object, or image, prioritize using the Content with Caption (7) or Picture with Caption (8) layouts. These layouts break up text, improve engagement, and are essential when a visual is implied or necessary for understanding. Avoid using these layouts gratuitously; use them when they truly add value, making the presentation outline more visually diverse and interactive.
7.  **Conciseness:** Keep text focused, but ensure important details are included as per the **COMPREHENSIVE COVERAGE** rule. Extract the core message for each point.
8.  **Visual Descriptions:** For layouts 7 & 8, provide clear, concise descriptions in `"object_description"` or `"picture_description"` to guide visual selection later.
9.  **Optional Speaker Notes:** Optionally include a `"notes"` (str) key on any slide object for speaker guidance.

**Example 1: IMAX Introduction (Using Layouts 0 & 1)**

**Input Document Snippet (Conceptual):** "IMAX is a high-resolution film format founded in 1967... known for large screens (1.43:1 or 1.90:1) and steep seating..."

**Output JSON:**
```json
{
  "slides": [
    {
      "layout_idx": 0, // Title Slide
      "title": "**IMAX**: The Ultimate Motion Picture Experience",
      "subtitle": "Understanding the Technology and Impact"
    },
    {
      "layout_idx": 1, // Title and Content
      "title": "What is IMAX?",
      "content": [
        "A proprietary system of **high-resolution** cameras, film, projectors, & theaters.",
        "Known for *very large screens* with tall aspect ratios (**1.43:1** or **1.90:1**).",
        "Features <u>steep stadium seating</u> for immersive viewing.",
        "Originally *Multiscreen Corporation* (founded 1967)."
      ],
      "notes": "Introduce the core concept and defining characteristics of IMAX."
    }
  ]
}
```

**Example 2: IMAX Evolution (Using Layouts 2 & 4)**

**Input Document Snippet (Conceptual):** "...Digital IMAX (2008) used dual 2K projectors, 1.90:1 ratio, lower cost but lower quality... IMAX with Laser (2014) uses dual 4K lasers, brighter, better contrast, Rec.2020 color, can show 1.43:1..."

**Output JSON:**
```json
{
  "slides": [
    // ... previous slides ...
    {
        "layout_idx": 2, // Section Header
        "section_title": "The Evolution: From Film to Digital & Laser",
        "section_description": "Adapting IMAX for broader adoption and enhanced quality"
    },
    {
        "layout_idx": 4, // Comparison
        "title": "IMAX Projection: Digital vs. Laser",
        "left_heading": "**Digital IMAX** (2008)",
        "right_heading": "**IMAX with Laser** (2014+)",
        "left_comparison_content": [
          "Dual **2K** Xenon Projectors",
          "Standard **1.90:1** Aspect Ratio",
          "*Lower* initial cost",
          "Concerns about quality ('*LieMAX*')",
          "Standard 6-channel sound"
        ],
        "right_comparison_content": [
          "Dual **4K** Laser Projectors",
          "Can project native **1.43:1**",
          "*Significantly brighter*",
          "<u>Higher contrast</u> & wider color (Rec. 2020)",
          "Enhanced **12-channel sound**"
        ],
        "notes": "Directly compare the key features and differences between the two main digital projection technologies."
    }
    // ... subsequent slides ...
  ]
}
```

**Example 3: IMAX Camera & Film (Using Layout 7 & 8)**

**Input Document Snippet (Conceptual):** "The IMAX film camera uses 65mm stock horizontally, 15 perforations per frame... extremely high resolution (~18K)... noisy due to vacuum system... limited 3-minute load... The film stock itself is 65mm wide..."

**Output JSON:**
```json
{
  "slides": [
    // ... previous slides ...
    {
        "layout_idx": 7, // Content with Caption
        "title": "Capturing the Image: The IMAX Film Camera",
        "caption_text": [
          "Uses **65mm** film stock horizontally.",
          "**15 perforations** per frame (*15/70* format).",
          "Achieves extreme resolution (est. **~18K**).",
          "<u>Challenges:</u>",
          "  - *Noisy* operation (vacuum system).",
          "  - Short (~3 min) film load.",
          "  - Shallow depth of field."
        ],
        "object_description": "Close-up photo or diagram of the unique IMAX 15/70 film camera, highlighting its size or horizontal film path.",
        "notes": "Explain the core technical aspects and trade-offs of the original IMAX film camera technology. Visual aid is key here."
    },
    {
      "layout_idx": 8, // Picture with Caption
      "picture_description": "High-resolution image of IMAX 15/70 film stock, clearly showing the perforations and width compared to standard film.",
      "caption_text": "IMAX 15/70 Film Stock: The Gold Standard in Resolution",
      "notes": "Discuss the unique characteristics of IMAX film stock, emphasizing its resolution and physical format. A visual helps here."
    }
    // ... concluding slide ...
  ]
}
```

---

Analyze the document content you have access to (the indexed PDF) and generate the JSON output following **ALL** the instructions above, especially the **CRITICAL INSTRUCTIONS FOR COVERAGE AND STRUCTURE** and the refined guidance on layout choice. Ensure a logical flow, cover all essential topics, generate sufficient slides, and include a strong concluding slide.

**Output JSON:**
"""

# --- Query Instruction for LLM ---
QUERY_INSTRUCTION = "Generate a detailed PowerPoint presentation outline in JSON format based on the document content, following the structure and examples provided in your instructions. Add images with the help of the image search engine. Use layout 8 and 3 with images as much as possible. Every presentation must have at least two picture with caption layouts. No placeholder or textbox must be empty in any of the slides. Add maximum bullet points in the slides to make the presentation more engaging."

