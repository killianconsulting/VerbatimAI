import difflib
import requests
from bs4 import BeautifulSoup
from docx import Document
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import sys
import webbrowser

# ------------------ Helper Functions ------------------

def get_document_url_pairs(docx_files):
    match_window = tk.Toplevel()
    match_window.title("Match DOCX Files to URLs")
    window_width = 1200  # Adjusted for full field sizes
    window_height = 400  # Reduced height only
    
    # Calculate screen center
    screen_width = match_window.winfo_screenwidth()
    screen_height = match_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    match_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    entries = []
    canvas = tk.Canvas(match_window)
    scrollbar = tk.Scrollbar(match_window, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    tk.Label(scroll_frame, text="Enter the URL that matches each DOCX file:", font=("Arial", 12, "bold")).pack(pady=10)
    for file in docx_files:
        frame = tk.Frame(scroll_frame)
        frame.pack(fill="x", padx=10, pady=5)
        tk.Label(frame, text=file, width=80, anchor="w").pack(side="left")  # Back to original width
        url_entry = tk.Entry(frame, width=100)  # Back to original width
        url_entry.pack(side="left", padx=5, fill="x", expand=True)
        entries.append((file, url_entry))

    matched_pairs = []

    def submit():
        for filename, entry in entries:
            url = entry.get().strip()
            if not url:
                messagebox.showerror("Missing URL", f"Please enter a URL for {filename}")
                return
            matched_pairs.append((filename, url))
        match_window.destroy()

    submit_btn = tk.Button(scroll_frame, text="Submit Matches", command=submit)
    submit_btn.pack(pady=20)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    match_window.grab_set()
    match_window.wait_window()
    return matched_pairs

# ------------------ Remaining Functions ------------------

def get_webpage_text(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, "html.parser")
        
        # Get title
        title = "Untitled Page"
        if soup.title and soup.title.string:
            title = soup.title.string.strip()
        
        # Get meta description
        meta_description = ""
        meta_desc_tag = soup.find("meta", attrs={"name": "description"})
        if meta_desc_tag and meta_desc_tag.get("content"):
            meta_description = meta_desc_tag["content"].strip()
        
        # Try different content containers
        content_containers = [
            soup.find("main"),
            soup.find("article"),
            soup.find("div", {"class": ["content", "main-content", "page-content"]}),
            soup.find("body")
        ]
        
        main = next((container for container in content_containers if container is not None), None)
        if not main:
            return "[ERROR: Could not find main content area]", title, meta_description
        
        # Extract clean paragraphs while preserving structure
        paragraphs = []
        
        # First, handle regular content
        for tag in main.find_all(["p", "li", "h1", "h2", "h3", "h4", "h5", "h6"]):
            # Skip empty tags
            if not tag.get_text(strip=True):
                continue
                
            # Skip if inside structured content section to avoid duplication
            if tag.find_parent(class_=lambda x: x and any(keyword in str(x).lower() for keyword in [
                'faq', 'accordion', 'expandable', 'collapse', 'toggle',
                'uagb-faq', 'uagb-container', 'wp-block-uagb'
            ])):
                continue
                
            # Create a copy to work with
            tag_copy = BeautifulSoup(str(tag), "html.parser")
            
            # Handle links by preserving their text
            for a in tag_copy.find_all('a'):
                if a.get_text(strip=True):
                    a.unwrap()
            
            # Get the complete text of the element
            text = tag_copy.get_text(" ", strip=True)
            if text and len(text) > 1:
                # For headings, check if they're visible and not hidden by CSS
                if tag.name.startswith('h'):
                    # Skip headings that are likely hidden
                    parent_style = tag.get('style', '') + ' '.join(parent.get('style', '') for parent in tag.parents if parent.get('style'))
                    if any(style in parent_style.lower() for style in ['display: none', 'visibility: hidden']):
                        continue
                    # Skip headings inside navigation, header, or footer
                    if tag.find_parent(['nav', 'header', 'footer']):
                        continue
                    # Skip headings that are part of a menu or navigation
                    if any('menu' in cls.lower() or 'nav' in cls.lower() for cls in tag.get('class', [])):
                        continue
                    paragraphs.append(f"<{tag.name}>{text}</{tag.name}>")
                else:
                    paragraphs.append(text)
        
        # Then, handle structured content sections
        structured_content_patterns = [
            # UAGB FAQ patterns
            {'class_': lambda x: x and any(c for c in str(x).split() if c.startswith('uagb-faq'))},
            {'class_': lambda x: x and any(c for c in str(x).split() if c.startswith('wp-block-uagb-faq'))},
            # Generic FAQ patterns
            {'class_': lambda x: x and any(keyword in str(x).lower() for keyword in ['faq', 'frequently-asked'])},
            # Accordion patterns
            {'class_': lambda x: x and any(keyword in str(x).lower() for keyword in ['accordion', 'expandable', 'collapse'])},
            # ARIA patterns
            {'role': 'tablist'},
            {'role': 'tab'},
            # Container patterns
            {'class_': lambda x: x and 'uagb-container-inner-blocks-wrap' in str(x)}
        ]
        
        # Find all structured content sections
        structured_sections = []
        for pattern in structured_content_patterns:
            sections = main.find_all(**pattern)
            structured_sections.extend(sections)
        
        # Remove duplicates while preserving order
        seen = set()
        structured_sections = [x for x in structured_sections if not (str(x) in seen or seen.add(str(x)))]
        
        # Process each structured section
        for section in structured_sections:
            # Try to find a section heading first
            section_heading = section.find(class_=lambda x: x and 'uagb-heading-text' in str(x))
            if section_heading and section_heading.get_text(strip=True):
                paragraphs.append(f"<h2>{section_heading.get_text(strip=True)}</h2>")
            
            # Find all question/answer pairs using multiple approaches
            qa_pairs = []
            
            # Method 1: UAGB FAQ structure
            questions = section.find_all(class_='uagb-question')
            for question in questions:
                # Get the FAQ item container
                faq_item = question.find_parent(class_=lambda x: x and 'uagb-faq-item' in str(x))
                if faq_item:
                    # Find the answer within this FAQ item
                    answer = faq_item.find(class_='uagb-faq-content')
                    if answer:
                        q_text = ' '.join(question.stripped_strings)
                        a_text = ' '.join(answer.stripped_strings)
                        if q_text and a_text:
                            qa_pairs.append((q_text, a_text))
            
            # Method 2: Generic FAQ/Accordion structure
            if not qa_pairs:
                questions = section.find_all(lambda tag: (
                    tag.name in ['dt', 'summary'] or
                    (tag.get('class') and any(c for c in tag.get('class', []) if any(keyword in c.lower() for keyword in ['question', 'header', 'title', 'summary']))) or
                    tag.get('role') == 'tab'
                ))
                
                for question in questions:
                    q_text = ' '.join(question.stripped_strings)
                    if not q_text:
                        continue
                    
                    # Try to find the corresponding answer
                    answer = None
                    
                    # Check for next sibling first
                    answer = question.find_next_sibling(lambda tag: (
                        tag.name == 'dd' or
                        (tag.get('class') and any(c for c in tag.get('class', []) if any(keyword in c.lower() for keyword in ['answer', 'content', 'panel', 'body']))) or
                        tag.get('role') == 'tabpanel'
                    ))
                    
                    # If no sibling found, try parent's next element
                    if not answer and question.parent:
                        answer = question.parent.find_next(lambda tag: (
                            tag.name == 'dd' or
                            (tag.get('class') and any(c for c in tag.get('class', []) if any(keyword in c.lower() for keyword in ['answer', 'content', 'panel', 'body']))) or
                            tag.get('role') == 'tabpanel'
                        ))
                    
                    if answer:
                        a_text = ' '.join(answer.stripped_strings)
                        if a_text:
                            qa_pairs.append((q_text, a_text))
            
            # Add all found Q&A pairs to paragraphs
            for q_text, a_text in qa_pairs:
                paragraphs.append(f"Q: {q_text}")
                paragraphs.append(f"A: {a_text}")
        
        if not paragraphs:
            return "[ERROR: No content found on page]", title, meta_description
            
        # Join paragraphs with double newlines to preserve structure
        raw_text = "\n\n".join(paragraphs)
        return raw_text, title, meta_description
        
    except requests.exceptions.RequestException as e:
        return f"[ERROR: Failed to fetch webpage: {str(e)}]", "Untitled Page", ""
    except Exception as e:
        return f"[ERROR: {str(e)}]", "Untitled Page", ""

def get_docx_text(path):
    doc = Document(path)
    paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip():
            # Preserve formatting for headings
            if p.style.name.startswith('Heading'):
                paragraphs.append(f"<h{p.style.name[-1]}>{p.text}</h{p.style.name[-1]}>")
            else:
                paragraphs.append(p.text)
    return "\n\n".join(paragraphs)

def normalize_text(text):
    # Only normalize whitespace and line breaks, preserve the rest
    text = re.sub(r"\r", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()

def normalize_html(text):
    # First, preserve link text by unwrapping anchor tags
    soup = BeautifulSoup(text, 'html.parser')
    for a in soup.find_all('a'):
        a.unwrap()
    text = str(soup)
    
    # Then proceed with other normalizations
    text = re.sub(r"<ul.*?>", "", text)
    text = re.sub(r"</ul.*?>", "", text)
    text = re.sub(r"<li.*?>", "• ", text)
    text = re.sub(r"</li.*?>", "", text)
    text = re.sub(r"</?(strong|b)>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"([a-zA-Z])\s*:\s+", r"\1: ", text)
    return text.strip()

def split_into_blocks(text):
    return [block.strip() for block in text.split("\n\n") if block.strip()]

def block_compare(draft, live, similarity_threshold=0.9):
    # Split into blocks while preserving paragraph structure
    draft_blocks = split_into_blocks(draft)
    live_blocks = split_into_blocks(live)
    
    # Find the first H1 in both draft and live content
    draft_h1_index = next((i for i, block in enumerate(draft_blocks) if block.startswith('<h1>')), -1)
    live_h1_index = next((i for i, block in enumerate(live_blocks) if block.startswith('<h1>')), -1)
    
    # Track total content and matched content for similarity calculation
    total_draft_length = sum(len(block) for block in draft_blocks[draft_h1_index:]) if draft_h1_index != -1 else sum(len(block) for block in draft_blocks)
    total_live_length = sum(len(block) for block in live_blocks[live_h1_index:]) if live_h1_index != -1 else sum(len(block) for block in live_blocks)
    matched_content_length = 0
    
    # Initialize aligned results list
    aligned = []
    matched_live = set()

    def is_heading(text):
        return any(text.startswith(f'<h{i}>') for i in range(1, 7))

    def get_content_type(text):
        if text.startswith('Page Name:'):
            return 'page_name'
        elif text.startswith('Internal Reference:'):
            return 'internal_ref'
        elif text.startswith('Page Link:'):
            return 'page_link'
        elif text.startswith('Meta Title:'):
            return 'meta_title'
        elif text.startswith('Meta Description:'):
            return 'meta_desc'
        elif is_heading(text):
            return 'heading'
        return 'content'

    def calculate_similarity(text1, text2):
        type1 = get_content_type(text1)
        type2 = get_content_type(text2)
        
        # If types don't match, they're not similar
        if type1 != type2:
            return 0.0
        
        # For headings and metadata, require exact matches after stripping HTML tags
        if type1 in ['heading', 'page_name', 'internal_ref', 'page_link', 'meta_title', 'meta_desc']:
            # Strip HTML tags for comparison
            clean1 = re.sub(r'<[^>]+>', '', text1)
            clean2 = re.sub(r'<[^>]+>', '', text2)
            
            # For page names and headings, strip the prefix
            if type1 in ['page_name', 'heading']:
                clean1 = clean1.replace('Page Name:', '').strip()
                clean2 = clean2.replace('Page Name:', '').strip()
            
            # Exact match required for these types
            return 1.0 if clean1 == clean2 else 0.0
        
        # For regular content, use sequence matcher with high threshold
        return difflib.SequenceMatcher(None, text1, text2).ratio()
    
    # First pass: try to match blocks with high similarity
    for i, db in enumerate(draft_blocks):
        best_match = None
        best_score = 0
        best_index = -1
        
        # Try to find the best matching block in live content
        for j, lb in enumerate(live_blocks):
            if lb in matched_live:
                continue
            
            score = calculate_similarity(db, lb)
            if score > best_score:
                best_score = score
                best_match = lb
                best_index = j
        
        # If we have a good match, use it
        if best_score >= similarity_threshold:
            matched_live.add(best_match)
            # Add any unmatched live blocks that come before this match
            for k in range(best_index):
                if live_blocks[k] not in matched_live:
                    # Check if this block has higher similarity with any upcoming draft blocks
                    future_match = False
                    for future_db in draft_blocks[i+1:]:
                        future_score = calculate_similarity(future_db, live_blocks[k])
                        if future_score >= similarity_threshold:
                            future_match = True
                            break
                    if not future_match:
                        aligned.append(("current", "", live_blocks[k]))
                        matched_live.add(live_blocks[k])
            
            aligned.append(("matched", db, best_match))
            matched_content_length += len(db) * best_score
        else:
            # If no good match, check for partial matches
            partial_matches = []
            partial_match_length = 0
            best_partial_index = -1
            
            # Skip partial matching for headings and metadata
            if get_content_type(db) not in ['heading', 'page_name', 'internal_ref', 'page_link', 'meta_title', 'meta_desc']:
                for j, lb in enumerate(live_blocks):
                    if lb in matched_live:
                        continue
                    # Skip partial matching if types don't match
                    if get_content_type(db) != get_content_type(lb):
                        continue
                    
                    # Split into sentences and check for partial matches
                    live_sentences = [s.strip() for s in lb.split('.') if s.strip()]
                    draft_sentences = [s.strip() for s in db.split('.') if s.strip()]
                    
                    sentence_matches = []
                    for ds in draft_sentences:
                        for ls in live_sentences:
                            match_score = difflib.SequenceMatcher(None, ds, ls).ratio()
                            if match_score > 0.8:  # Lower threshold for partial matches
                                sentence_matches.append((ds, ls, match_score))
                                partial_match_length += len(ds) * match_score
                    
                    if sentence_matches:
                        partial_matches.extend(sentence_matches)
                        if best_partial_index == -1:
                            best_partial_index = j
            
            if partial_matches:
                # Add any unmatched live blocks that come before this partial match
                for k in range(best_partial_index):
                    if live_blocks[k] not in matched_live:
                        aligned.append(("current", "", live_blocks[k]))
                        matched_live.add(live_blocks[k])
                
                # Combine partial matches
                combined_live = " ".join(m[1] for m in partial_matches)
                matched_live.add(combined_live)
                aligned.append(("matched", db, combined_live))
                matched_content_length += partial_match_length
            else:
                aligned.append(("missing", db, ""))
    
    # Add any remaining unmatched live blocks at their relative positions
    for i, lb in enumerate(live_blocks):
        if lb not in matched_live:
            # Try to find the best position based on similarity with surrounding content
            best_pos = len(aligned)
            best_context_score = 0
            
            for pos in range(len(aligned) + 1):
                context_score = 0
                # Check similarity with previous block
                if pos > 0:
                    prev_content = aligned[pos-1][1] or aligned[pos-1][2]  # Use draft or live content
                    if prev_content:
                        context_score += calculate_similarity(prev_content, lb)
                
                # Check similarity with next block
                if pos < len(aligned):
                    next_content = aligned[pos][1] or aligned[pos][2]  # Use draft or live content
                    if next_content:
                        context_score += calculate_similarity(next_content, lb)
                
                if context_score > best_context_score:
                    best_context_score = context_score
                    best_pos = pos
            
            # Insert at best position
            aligned.insert(best_pos, ("current", "", lb))
    
    # Calculate similarity score
    if total_draft_length == 0 or total_live_length == 0:
        similarity = 0.0
    else:
        draft_similarity = matched_content_length / total_draft_length
        live_similarity = matched_content_length / total_live_length
        matched_blocks = sum(1 for tag, _, _ in aligned if tag == "matched")
        total_blocks = len(aligned)
        block_similarity = matched_blocks / total_blocks if total_blocks > 0 else 0
        similarity = max(draft_similarity, live_similarity) * 0.7 + block_similarity * 0.3
    
    return aligned, similarity

def format_result_as_html(docx_file, url, title, meta_desc, similarity, results):
    # Add title and color key
    report = """
    <div style='margin-bottom: 30px;'>
        <h1 style='font-family: Roboto, Arial, sans-serif; font-size: 32px; font-weight: bold; margin: 0 0 20px 0;'>VerbatimAI</h1>
        <div style='background: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 20px;'>
            <strong>Color Key:</strong>
            <ul style='margin: 10px 0; padding-left: 20px;'>
                <li><span style='color: #28a745;'>Green</span> - Content matches between draft and live site</li>
                <li><span style='color: #dc3545;'>Red</span> - Content in draft but missing from live site</li>
                <li><span style='color: #007bff;'>Blue</span> - Content on live site but not in draft</li>
            </ul>
        </div>
    </div>
    """
    
    report += f"<h2>{docx_file} vs <a href='{url}'>{url}</a></h2>"
    report += f"<p><strong>Page Title:</strong> {title}</p>"
    report += f"<p><strong>Meta Description:</strong> {meta_desc}</p>"
    report += f"<p><strong>Similarity Score:</strong> {similarity:.2%}</p>"
    
    if similarity > 0.95:
        report += "<p style='color: #28a745;'>✅ Content is mostly identical.</p>"
    elif similarity > 0.75:
        report += "<p style='color: #ffc107;'>⚠️ Content has minor differences.</p>"
    else:
        report += "<p style='color: #dc3545;'>❌ Content is significantly different.</p>"

    # Add CSS for flex container and content rows
    report += """
    <style>
        .content-container {
            display: flex;
            margin-top: 20px;
            gap: 20px;
            align-items: stretch;
        }
        .column {
            flex: 1;
            display: flex;
            flex-direction: column;
        }
        .content-blocks {
            flex: 1;
            display: flex;
            flex-direction: column;
            gap: 10px;  /* Add gap between blocks */
        }
        .content-row {
            display: flex;
            width: 100%;
        }
        .content-block {
            width: 100%;
            padding: 15px;  /* Increased padding */
            box-sizing: border-box;
            margin: 0;
            line-height: 1.6;  /* Slightly increased line height */
            border-radius: 4px;  /* Rounded corners */
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);  /* Subtle shadow */
        }
        .matched-content {
            background-color: #e8f5e9;
            border: 1px solid #c8e6c9;  /* Subtle border */
        }
        .missing-content {
            background-color: #ffebee;
            border: 1px solid #ffcdd2;  /* Subtle border */
        }
        .current-content {
            background-color: #e3f2fd;
            border: 1px solid #bbdefb;  /* Subtle border */
        }
        .placeholder {
            border: 1px dashed #ddd;
            display: flex;
            align-items: center;
            justify-content: center;
            font-style: italic;
            color: #666;
            text-align: center;
            padding: 20px;
            background-color: #fafafa;  /* Light background */
        }
        .content-pair {
            display: flex;
            min-height: fit-content;
        }
    </style>
    """

    # Start the two-column layout
    report += "<div class='content-container'>"
    
    # Draft content column
    report += "<div class='column'>"
    report += "<h3>Draft Content</h3>"
    report += "<div class='content-blocks'>"
    
    # Live content column
    live_column = "<div class='column'>"
    live_column += "<h3>Live Content</h3>"
    live_column += "<div class='content-blocks'>"

    # Process results in their original order
    for tag, draft, live in results:
        # Calculate content length and approximate line count
        if tag == "matched":
            content_length = max(len(draft), len(live))
            line_count = max(draft.count('\n') + 1, live.count('\n') + 1)
            min_height = max(50, (content_length // 50 + line_count) * 24)  # 24px per line of text
            report += f"<div class='content-block matched-content' style='min-height: {min_height}px'>{draft}</div>"
            live_column += f"<div class='content-block matched-content' style='min-height: {min_height}px'>{live}</div>"
        elif tag == "missing":
            content_length = len(draft)
            line_count = draft.count('\n') + 1
            min_height = max(50, (content_length // 50 + line_count) * 24)
            report += f"<div class='content-block missing-content' style='min-height: {min_height}px'>{draft}</div>"
            live_column += f"<div class='content-block placeholder' style='min-height: {min_height}px'><em>Content missing from live site</em></div>"
        elif tag == "current":
            content_length = len(live)
            line_count = live.count('\n') + 1
            min_height = max(50, (content_length // 50 + line_count) * 24)
            report += f"<div class='content-block placeholder' style='min-height: {min_height}px'><em>Content not in draft</em></div>"
            live_column += f"<div class='content-block current-content' style='min-height: {min_height}px'>{live}</div>"

    # Close the columns
    report += "</div></div>"  # Close draft column
    live_column += "</div></div>"  # Close live column
    
    # Add the live column and close the container
    report += live_column + "</div>"
    
    return report

def format_result_as_markdown(docx_file, url, title, meta_desc, similarity, results):
    report = f"## {docx_file} vs {url}\n"
    report += f"**Page Title**: {title}\n\n"
    report += f"**Meta Description**: {meta_desc}\n\n"
    report += f"**Similarity Score**: `{similarity:.2%}`\n\n"
    if similarity > 0.95:
        report += "✅ Content is mostly identical.\n\n"
    elif similarity > 0.75:
        report += "⚠️ Content has minor differences.\n\n"
    else:
        report += "❌ Content is significantly different.\n\n"
    report += "### Differences\n"
    for tag, draft, live in results:
        if tag == "matched":
            report += f"✅ MATCHED: {draft}\n"
        elif tag == "missing":
            report += f"🟥 MISSING: {draft}\n"
            if live:
                report += f"🟩 CURRENT: {live}\n"
        elif tag == "current":
            report += f"🟩 CURRENT: {live}\n"
    report += "\n"
    return report

# ------------------ Main Comparison Logic ------------------

def run_batch_comparison():
    folder = filedialog.askdirectory(title="Select Folder Containing Draft DOCX Files")
    if not folder:
        return
    docx_files = sorted([f for f in os.listdir(folder) if f.endswith(".docx")])
    if not docx_files:
        messagebox.showerror("Error", "No .docx files found in the selected folder.")
        return
    matches = get_document_url_pairs(docx_files)
    if not matches:
        return
    total = len(matches)
    progress_bar["maximum"] = total
    progress_bar["value"] = 0
    report_md = "# Batch Comparison Report\n\n"
    summary = []
    for i, (docx_file, url) in enumerate(matches, start=1):
        full_path = os.path.join(folder, docx_file)
        try:
            draft_text = normalize_text(get_docx_text(full_path))
            live_text, title, meta_desc = get_webpage_text(url)
            live_text = normalize_text(live_text)
            if "[ERROR" in live_text:
                report_md += f"## {docx_file} vs {url}\n❌ {live_text}\n\n"
                summary.append(f"❌ {url}: Error")
                continue
            
            # Get both alignment results and similarity score from block_compare
            diff, similarity = block_compare(draft_text, live_text)
            
            html_report = format_result_as_html(docx_file, url, title, meta_desc, similarity, diff)
            markdown_report = format_result_as_markdown(docx_file, url, title, meta_desc, similarity, diff)

            html_file_path = os.path.join(folder, f"report_{i}_{os.path.splitext(docx_file)[0]}.html")
            with open(html_file_path, "w", encoding="utf-8") as f:
                f.write(f"""<html>
                    <head>
                        <meta charset='UTF-8'>
                        <title>VerbatimAI - Comparison Report</title>
                        <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">
                        <style>
                            body {{ 
                                font-family: Roboto, Arial, sans-serif; 
                                margin: 20px;
                                line-height: 1.6;
                            }}
                            h1 {{ 
                                font-size: 32px; 
                                font-weight: bold; 
                                margin: 0 0 20px 0;
                            }}
                            .color-key {{
                                background: #f5f5f5;
                                padding: 15px;
                                border-radius: 5px;
                                margin-bottom: 20px;
                            }}
                            .color-key ul {{
                                margin: 10px 0;
                                padding-left: 20px;
                            }}
                            .matched {{
                                background-color: #e8f5e9;
                            }}
                            .missing {{
                                background-color: #ffebee;
                            }}
                            .current {{
                                background-color: #e3f2fd;
                            }}
                            .placeholder {{
                                border: 1px dashed #ddd;
                                color: #666;
                                font-style: italic;
                            }}
                        </style>
                    </head>
                    <body>{html_report}</body>
                </html>""")
            if i == total:
                # Open the folder instead of the HTML file
                os.startfile(folder)

            report_md += markdown_report
            summary.append(f"{url} → Similarity: {similarity:.2%}")

        except Exception as e:
            report_md += f"## {docx_file} vs {url}\n❌ Error: {str(e)}\n\n"
            summary.append(f"❌ {url}: Error")
        progress_bar["value"] = i
        root.update_idletasks()
    md_path = os.path.join(folder, "comparison_report.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(report_md)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, "Reports saved.\n\n" + "\n".join(summary))
    messagebox.showinfo("Done", f"✅ Batch comparison complete.\nMarkdown saved to:\n{md_path}\nHTML reports saved alongside each docx.")
    progress_bar["value"] = 0

# ------------------ GUI Setup ------------------

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("VerbatimAI")

    # Set the window icon
    try:
        icon_path = resource_path('verbatim.ico')
        root.iconbitmap(default=icon_path)
    except Exception as e:
        print(f"Error loading icon: {e}")
        pass

    # Create main frame with padding
    frame = tk.Frame(root)
    frame.pack(padx=20, pady=20)

    # Add logo
    try:
        logo_path = resource_path('smbteam-logo.png')
        logo_image = tk.PhotoImage(file=logo_path)
        # Resize the image to a reasonable size
        logo_image = logo_image.subsample(2, 2)  # Adjust these values if needed
        logo_label = tk.Label(frame, image=logo_image)
        logo_label.image = logo_image  # Keep a reference to prevent garbage collection
        logo_label.pack(pady=(0, 10))
    except Exception as e:
        print(f"Error loading logo: {e}")
        pass

    # Add title text with Roboto font
    title_label = tk.Label(frame, text="VerbatimAI", font=("Roboto", 32, "bold"))
    title_label.pack(pady=(0, 20))

    # Add the main button
    button = tk.Button(frame, text="Start AutoCompare", command=run_batch_comparison)
    button.pack(pady=5)

    # Add progress bar
    progress_bar = ttk.Progressbar(frame, orient="horizontal", length=600, mode="determinate")
    progress_bar.pack(pady=5)

    # Add text area
    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=10)
    text_area.pack(padx=10, pady=10)

    root.mainloop() 