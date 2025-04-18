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
import shutil
import urllib.parse
from urllib.robotparser import RobotFileParser
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# Try to import tkinterdnd2, fall back to regular tkinter if not available
try:
    import tkinterdnd2 as tkdnd
    USE_DND = True
except ImportError:
    USE_DND = False

# ------------------ Helper Functions ------------------

def get_document_url_pairs(docx_files, parent_window):
    match_window = tk.Toplevel(parent_window)
    match_window.title("Match DOCX Files to URLs")
    window_width = 1200
    
    # Calculate initial height based on number of documents
    base_height = 200  # Base height for controls and padding
    entry_height = 30  # Height per entry
    window_height = base_height + (len(docx_files) * entry_height)
    # Ensure minimum height
    window_height = max(window_height, 300)
    # Ensure maximum height (80% of screen height)
    max_height = int(parent_window.winfo_screenheight() * 0.8)
    window_height = min(window_height, max_height)
    
    # Calculate screen center
    screen_width = match_window.winfo_screenwidth()
    screen_height = match_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    match_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Make window modal
    match_window.transient(parent_window)
    match_window.grab_set()
    
    # Check if dark mode is enabled
    current_settings = load_settings()
    is_dark_mode = current_settings.get('dark_mode', 'false').lower() == 'true'
    
    # Apply theme colors
    bg_color = '#1e1e1e' if is_dark_mode else '#f0f0f0'
    fg_color = '#ffffff' if is_dark_mode else '#000000'
    entry_bg = '#2d2d2d' if is_dark_mode else '#ffffff'
    button_bg = '#404040' if is_dark_mode else '#e0e0e0'
    button_fg = '#000000'  # Always black for better readability
    
    # Configure ttk styles for dark mode
    style = ttk.Style()
    if is_dark_mode:
        style.configure('url.TEntry', 
                       fieldbackground='#2d2d2d',
                       background='#2d2d2d',
                       foreground='#000000',
                       insertcolor='#ffffff',
                       selectbackground='#404040',
                       selectforeground='#ffffff')
        
        style.configure('url.TButton',
                       background='#404040',
                       foreground='#000000',
                       bordercolor='#505050',
                       lightcolor='#404040',
                       darkcolor='#2d2d2d',
                       relief='raised')
        
        style.map('url.TButton',
                 background=[('active', '#505050'), ('pressed', '#303030')],
                 foreground=[('active', '#000000'), ('pressed', '#000000')])

        # Add styles for the URL matching window elements
        style.configure('url.TLabel',
                       background='#1e1e1e',
                       foreground='#ffffff')
        
        style.configure('url.TFrame',
                       background='#1e1e1e')
        
        style.configure('url.TRadiobutton',
                       background='#1e1e1e',
                       foreground='#ffffff',
                       selectcolor='#404040')
    else:
        style.configure('url.TEntry',
                       fieldbackground='#ffffff',
                       background='#ffffff',
                       foreground='#000000',
                       insertcolor='#000000',
                       selectbackground='#0078d7',
                       selectforeground='#ffffff')
        
        style.configure('url.TButton',
                       background='#f0f0f0',
                       foreground='#000000')
        
        style.map('url.TButton',
                 background=[('active', '#e0e0e0'), ('pressed', '#cccccc')],
                 foreground=[('active', '#000000'), ('pressed', '#000000')])

        # Light mode styles for URL matching window
        style.configure('url.TLabel',
                       background='#f0f0f0',
                       foreground='#000000')
        
        style.configure('url.TFrame',
                       background='#f0f0f0')
        
        style.configure('url.TRadiobutton',
                       background='#f0f0f0',
                       foreground='#000000',
                       selectcolor='#ffffff')

    # Apply theme colors to the match window
    match_window.configure(bg=bg_color)
    
    # Create main frame
    main_frame = ttk.Frame(match_window, style='url.TFrame')
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Add mode selection
    mode_frame = ttk.Frame(main_frame, style='url.TFrame')
    mode_frame.pack(fill="x", pady=(0, 10))
    
    mode_var = tk.StringVar(value="auto")
    
    def on_mode_change():
        if mode_var.get() == "auto":
            base_url_frame.pack(fill="x", pady=5)
            manual_frame.pack_forget()
        else:
            base_url_frame.pack_forget()
            manual_frame.pack(fill="both", expand=True)
    
    # Add radio buttons with theme
    ttk.Radiobutton(mode_frame, text="Automatic URL Matching", variable=mode_var, value="auto", 
                   command=on_mode_change, style='url.TRadiobutton').pack(side="left", padx=5)
    ttk.Radiobutton(mode_frame, text="Manual URL Entry", variable=mode_var, value="manual", 
                   command=on_mode_change, style='url.TRadiobutton').pack(side="left", padx=5)
    
    # Base URL entry for automatic mode
    base_url_frame = ttk.Frame(main_frame, style='url.TFrame')
    
    # Add description for automatic mode
    auto_description = ttk.Label(
        base_url_frame,
        text="Automatic URL Matching will crawl the website starting from the domain URL you provide. " +
             "It will find and analyze all pages on the site, compare each document with the found pages, " +
             "and automatically match your documents to their best corresponding URLs.",
        font=("Arial", 12, "bold"),
        style='url.TLabel',
        justify='left',
        wraplength=550
    )
    auto_description.pack(pady=5)
    
    base_url_label = ttk.Label(base_url_frame, text="Enter Domain:", style='url.TLabel')
    base_url_label.pack(side="left", padx=5)
    base_url_entry = ttk.Entry(base_url_frame, width=100, style='url.TEntry')
    base_url_entry.pack(side="left", fill="x", expand=True, padx=5)
    
    # Manual URL entry frame
    manual_frame = ttk.Frame(main_frame, style='url.TFrame')
    
    # Create a container frame for the scrollable content
    scroll_container = ttk.Frame(manual_frame, style='url.TFrame')
    scroll_container.pack(fill="both", expand=True)
    
    entries = []
    canvas = tk.Canvas(scroll_container, bg=bg_color, highlightthickness=0)
    scrollbar = tk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas, style='url.TFrame')

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Add instructions with theme-aware colors
    instructions = ttk.Label(scroll_frame, 
                           text="Enter the URL that matches each DOCX file.\nTip: You can paste multiple URLs at once!", 
                           font=("Arial", 12, "bold"),
                           style='url.TLabel')
    instructions.pack(pady=5)

    # Create a frame for the paste button
    paste_frame = ttk.Frame(scroll_frame, style='url.TFrame')
    paste_frame.pack(fill="x", padx=10, pady=5)

    def paste_urls():
        try:
            # Get clipboard content
            clipboard = match_window.clipboard_get()
            # Split into lines and clean up
            urls = [url.strip() for url in clipboard.splitlines() if url.strip()]
            
            # Fill as many entry fields as we have URLs
            for i, url in enumerate(urls):
                if i < len(entries):
                    entries[i][1].delete(0, tk.END)
                    entries[i][1].insert(0, url)
        except tk.TclError:
            messagebox.showwarning("Clipboard Empty", "No text found in clipboard.")
        except Exception as e:
            messagebox.showerror("Error", f"Error pasting URLs: {str(e)}")

    paste_btn = ttk.Button(paste_frame, text="Paste URLs from Clipboard", command=paste_urls, style='url.TButton')
    paste_btn.pack(side="right", padx=10)

    # Add tooltip for paste button with theme-aware colors
    paste_tooltip = ttk.Label(paste_frame, 
                            text="Paste a list of URLs (one per line) or click here to auto-fill multiple copied urls",
                            style='url.TLabel')
    paste_tooltip.pack(side="right", padx=5)

    # Create entry fields
    for file in docx_files:
        frame = ttk.Frame(scroll_frame, style='url.TFrame')
        frame.pack(fill="x", padx=10, pady=2)
        
        # File label with fixed width and theme-aware colors
        file_label = ttk.Label(frame, text=os.path.basename(file), width=80, anchor="w", style='url.TLabel')
        file_label.pack(side="left")
        
        # URL entry with theme-aware colors
        url_entry = ttk.Entry(frame, width=100, style='url.TEntry')
        url_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Enable Ctrl+V for pasting in entry
        url_entry.bind('<Control-v>', lambda e: 'break')  # Prevent default paste
        url_entry.bind('<Control-V>', lambda e: 'break')  # Prevent default paste
        
        entries.append((file, url_entry))

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    matched_pairs = []

    def submit():
        nonlocal matched_pairs
        if mode_var.get() == "auto":
            base_url = base_url_entry.get().strip()
            if not base_url:
                messagebox.showerror("Missing URL", "Please enter a base URL")
                return
            matches = auto_match_documents(docx_files, base_url, match_window)
            if matches:
                matched_pairs = matches
                match_window.grab_release()
                match_window.destroy()
        else:
            for filename, entry in entries:
                url = entry.get().strip()
                if not url:
                    messagebox.showerror("Missing URL", f"Please enter a URL for {os.path.basename(filename)}")
                    return
                matched_pairs.append((filename, url))
            match_window.grab_release()
            match_window.destroy()

    # Create button frame at the bottom
    button_frame = ttk.Frame(main_frame, style='url.TFrame')
    button_frame.pack(fill="x", side="bottom", pady=10)
    
    # Create submit button in button frame
    submit_btn = ttk.Button(button_frame, text="Start AutoCompare", command=submit, style='url.TButton')
    submit_btn.pack()
    
    # Bind Enter key to submit
    match_window.bind('<Return>', lambda e: submit())
    base_url_entry.bind('<Return>', lambda e: submit())
    
    # Initialize the correct mode display
    on_mode_change()
    
    # Wait for window to be destroyed
    match_window.wait_window()
    return matched_pairs

# ------------------ Remaining Functions ------------------

def get_webpage_text(url):
    try:
        # Configure session with proper settings
        session = requests.Session()
        session.verify = False  # Ignore SSL certificate verification
        session.max_redirects = 5  # Allow up to 5 redirects
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Cache-Control': 'max-age=0'
        }

        # Suppress only the single InsecureRequestWarning from urllib3
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        
        # First try HTTPS
        try:
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
            response = session.get(url, headers=headers, timeout=30, allow_redirects=True)
            response.raise_for_status()
        except requests.exceptions.SSLError:
            # If HTTPS fails, try HTTP
            if url.startswith('https://'):
                url = 'http://' + url[8:]
            response = session.get(url, headers=headers, timeout=30, allow_redirects=True)
            response.raise_for_status()
        
        # Check content type
        content_type = response.headers.get('content-type', '').lower()
        if 'text/html' not in content_type:
            return f"[ERROR: Invalid content type: {content_type}]", "Untitled Page", ""
        
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
            soup.find(id=lambda x: x and any(word in str(x).lower() for word in ['content', 'main', 'article'])),
            soup.find(class_=lambda x: x and any(word in str(x).lower() for word in ['content', 'main-content', 'page-content', 'article'])),
            soup.find("div", {"class": ["content", "main-content", "page-content", "article-content"]}),
            soup.find("body")
        ]
        
        main = next((container for container in content_containers if container is not None), None)
        if not main:
            return "[ERROR: Could not find main content area]", title, meta_description
        
        # Remove unwanted elements
        for element in main.find_all(['script', 'style', 'iframe', 'noscript', 'header', 'footer', 'nav']):
            element.decompose()
        
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
        
    except requests.exceptions.SSLError as e:
        return f"[ERROR: SSL Certificate verification failed: {str(e)}]", "Untitled Page", ""
    except requests.exceptions.ConnectionError as e:
        return f"[ERROR: Failed to connect to server: {str(e)}]", "Untitled Page", ""
    except requests.exceptions.Timeout as e:
        return f"[ERROR: Request timed out: {str(e)}]", "Untitled Page", ""
    except requests.exceptions.TooManyRedirects as e:
        return f"[ERROR: Too many redirects: {str(e)}]", "Untitled Page", ""
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
    <div class='report-container'>
        <div class='header'>
            <h1>Verbatim AI Content Comparison</h1>
            <div class='color-key'>
            <strong>Color Key:</strong>
                <ul>
                    <li><span class='matched-text'>Green</span> - Content matches between draft and live site</li>
                    <li><span class='missing-text'>Red</span> - Content in draft but missing from live site</li>
                    <li><span class='current-text'>Blue</span> - Content on live site but not in draft</li>
            </ul>
        </div>
    </div>

        <div class='page-info'>
            <h2>Document Comparison</h2>
            <p><strong>Document:</strong> {docx_file}</p>
            <p><strong>URL:</strong> <a href='{url}' target='_blank'>{url}</a></p>
            <p><strong>Page Title:</strong> {title}</p>
            <p><strong>Meta Description:</strong> {meta_desc}</p>
            <p><strong>Similarity Score:</strong> {similarity:.2%}</p>
            {similarity_indicator}
        </div>

        <div class='content-comparison'>
            <div class='column-headers'>
                <h3>Draft Content</h3>
                <h3>Live Content</h3>
            </div>
            <div class='content-container'>
    """.format(
        docx_file=docx_file,
        url=url,
        title=title,
        meta_desc=meta_desc,
        similarity=similarity,
        similarity_indicator="""
            <p class='similarity-high' style='color: #28a745;'>✅ Content is mostly identical.</p>
        """ if similarity > 0.95 else """
            <p class='similarity-medium' style='color: #ffc107;'>⚠️ Content has minor differences.</p>
        """ if similarity > 0.75 else """
            <p class='similarity-low' style='color: #dc3545;'>❌ Content is significantly different.</p>
        """
    )

    # Process each content block
    for tag, draft, live in results:
        report += "<div class='content-row'>"
        
        # Draft column
        if tag == "matched":
            report += f"<div class='content-block matched-content'>{draft}</div>"
        elif tag == "missing":
            report += f"<div class='content-block missing-content'>{draft}</div>"
        else:  # current
            report += "<div class='content-block placeholder'><em>Content not in draft</em></div>"
        
        # Live column
        if tag == "matched":
            report += f"<div class='content-block matched-content'>{live}</div>"
        elif tag == "missing":
            report += "<div class='content-block placeholder'><em>Content missing from live site</em></div>"
        else:  # current
            report += f"<div class='content-block current-content'>{live}</div>"
        
        report += "</div>"

    # Close containers and add CSS
    report += """
            </div>
        </div>
    </div>
    <style>
        body {
            font-family: 'Roboto', Arial, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            background-color: #f8f9fa;
        }
        .report-container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .header {
            margin-bottom: 30px;
        }
        .header h1 {
            font-size: 32px;
            margin: 0 0 20px 0;
            color: #2c3e50;
        }
        .color-key {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        .color-key ul {
            list-style-type: none;
            padding-left: 0;
            margin: 10px 0;
        }
        .color-key li {
            margin: 5px 0;
        }
        .matched-text { color: #28a745; }
        .missing-text { color: #dc3545; }
        .current-text { color: #007bff; }
        .page-info {
            margin-bottom: 30px;
            padding: 20px;
            background-color: #f8f9fa;
            border-radius: 5px;
        }
        .page-info h2 {
            margin-top: 0;
            color: #2c3e50;
        }
        .page-info p {
            margin: 10px 0;
        }
        .column-headers {
            display: flex;
            justify-content: space-between;
            margin-bottom: 15px;
        }
        .column-headers h3 {
            flex: 1;
            margin: 0;
            padding: 10px;
            background-color: #f8f9fa;
            border-radius: 5px;
            text-align: center;
            color: #2c3e50;
        }
        .content-container {
            display: flex;
            flex-direction: column;
            gap: 15px;
        }
        .content-row {
            display: flex;
            gap: 15px;
            min-height: fit-content;
        }
        .content-block {
            flex: 1;
            padding: 15px;
            border-radius: 5px;
            white-space: pre-wrap;
            word-break: break-word;
            min-height: 50px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .matched-content {
            background-color: #e8f5e9;
            border: 1px solid #c8e6c9;
        }
        .missing-content {
            background-color: #ffebee;
            border: 1px solid #ffcdd2;
        }
        .current-content {
            background-color: #e3f2fd;
            border: 1px solid #bbdefb;
        }
        .placeholder {
            background-color: #f8f9fa;
            border: 1px dashed #dee2e6;
            display: flex;
            align-items: center;
            justify-content: center;
            font-style: italic;
            color: #6c757d;
        }
        a {
            color: #007bff;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
    </style>
    """
    
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

def handle_drop(event):
    """Handle dropped files or folders"""
    data = event.data
    if not data:
        return
        
    # Remove curly braces if present
    data = data.strip('{}')
    
    # Split multiple files and handle paths with spaces
    paths = []
    current_path = ""
    in_quotes = False
    
    for char in data:
        if char == '"':
            in_quotes = not in_quotes
        elif char == ' ' and not in_quotes:
            if current_path:
                paths.append(current_path)
                current_path = ""
        else:
            current_path += char
    
    if current_path:
        paths.append(current_path)
    
    # Collect all DOCX files
    docx_files = []
    for path in paths:
        # Remove quotes if present
        path = path.strip('"')
        
        if os.path.isdir(path):
            # If it's a directory, find all DOCX files in it
            docx_files.extend([
                os.path.join(path, f) 
                for f in os.listdir(path) 
                if f.lower().endswith('.docx')
            ])
        elif path.lower().endswith('.docx'):
            # If it's a DOCX file, add it directly
            docx_files.append(path)
    
    if not docx_files:
        messagebox.showerror("Error", "No .docx files found in the dropped items.")
        return
    
    # Create a results folder
    first_file_dir = os.path.dirname(docx_files[0])
    results_folder = os.path.join(first_file_dir, "VerbatimAI_Results")
    if not os.path.exists(results_folder):
        os.makedirs(results_folder)
    
    # Process the files
    process_files(results_folder, docx_files)

def run_batch_comparison(folder=None):
    """Run comparison for files selected through folder, multiple files, or single file"""
    if not folder:
        # Create a custom dialog for selection method
        selection_window = tk.Toplevel(root)
        selection_window.title("Select Upload Method")
        selection_window.geometry("400x200")
        
        # Make window modal
        selection_window.transient(root)
        selection_window.grab_set()
        
        # Center the window
        window_width = 400
        window_height = 200
        screen_width = selection_window.winfo_screenwidth()
        screen_height = selection_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        selection_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Check if dark mode is enabled
        current_settings = load_settings()
        is_dark_mode = current_settings.get('dark_mode', 'false').lower() == 'true'
        
        # Apply theme colors
        bg_color = '#1e1e1e' if is_dark_mode else '#f0f0f0'
        fg_color = '#ffffff' if is_dark_mode else '#000000'
        button_bg = '#404040' if is_dark_mode else '#e0e0e0'
        
        # Configure window colors
        selection_window.configure(bg=bg_color)
        
        # Configure ttk styles for the selection window
        style = ttk.Style()
        if is_dark_mode:
            style.configure('Select.TFrame', 
                          background='#1e1e1e')
            
            style.configure('Select.TButton',
                          background='#404040',
                          foreground='#000000',
                          bordercolor='#505050',
                          lightcolor='#404040')
            
            style.map('Select.TButton',
                     background=[('active', '#505050'), ('pressed', '#303030')],
                     foreground=[('active', '#000000'), ('pressed', '#000000')])
            
            style.configure('Select.TLabel',
                          background='#1e1e1e',
                          foreground='#ffffff')
        else:
            style.configure('Select.TFrame', 
                          background='#f0f0f0')
            
            style.configure('Select.TButton',
                          background='#f0f0f0',
                          foreground='#000000')
            
            style.map('Select.TButton',
                     background=[('active', '#e0e0e0'), ('pressed', '#cccccc')],
                     foreground=[('active', '#000000'), ('pressed', '#000000')])
            
            style.configure('Select.TLabel',
                          background='#f0f0f0',
                          foreground='#000000')

        def select_folder():
            selection_window.destroy()
            folder_path = filedialog.askdirectory(title="Select Folder Containing Draft DOCX Files")
            if folder_path:
                process_files(folder_path)

        def select_files():
            selection_window.destroy()
            files = filedialog.askopenfilenames(
                title="Select DOCX Files",
                filetypes=[("DOCX files", "*.docx"), ("All files", "*.*")]
            )
            if files:
                # Create temporary folder for selected files
                temp_folder = os.path.join(os.path.dirname(files[0]), "VerbatimAI_Results")
                if not os.path.exists(temp_folder):
                    os.makedirs(temp_folder)
                
                # Process selected files
                process_files(temp_folder, files)

        # Add descriptive labels and buttons with theme-aware styling
        title_label = tk.Label(
            selection_window,
            text="Choose how you want to upload documents:",
            font=("Roboto", 12),
            bg=bg_color,
            fg=fg_color
        )
        title_label.pack(pady=20)

        button_frame = ttk.Frame(selection_window, style='Select.TFrame')
        button_frame.pack(fill='x', padx=20)

        folder_button = ttk.Button(
            button_frame,
            text="Select Folder",
            command=select_folder,
            style='Select.TButton'
        )
        folder_button.pack(fill='x', pady=5)
        
        folder_label = ttk.Label(
            button_frame,
            text="Upload an entire folder of DOCX files",
            foreground='#a0a0a0',
            style='Select.TLabel'
        )
        folder_label.pack()

        files_button = ttk.Button(
            button_frame,
            text="Select Files",
            command=select_files,
            style='Select.TButton'
        )
        files_button.pack(fill='x', pady=(15,5))
        
        files_label = ttk.Label(
            button_frame,
            text="Choose one or multiple DOCX files",
            foreground='#a0a0a0',
            style='Select.TLabel'
        )
        files_label.pack()

        return

    process_files(folder)

def process_files(folder, specific_files=None):
    """Process the selected files for comparison"""
    # Get DOCX files based on selection method
    if specific_files:
        # Store original file paths and use them directly
        docx_files = specific_files
    else:
        # Get full paths for files in the folder
        docx_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.docx')]

    if not docx_files:
        messagebox.showerror("Error", "No .docx files found in the selected location.")
        return

    # Create a results folder
    first_file_dir = os.path.dirname(docx_files[0])
    results_folder = os.path.join(first_file_dir, "VerbatimAI_Results")
    if not os.path.exists(results_folder):
        os.makedirs(results_folder)

    # Create URL matching window
    matches = get_document_url_pairs(docx_files, root)
    if not matches:
        return

    # Disable the main window's drop target while processing
    if USE_DND:
        drop_target.drop_target_unregister()

    try:
        total = len(matches)
        progress_bar["maximum"] = total
        progress_bar["value"] = 0
        report_md = "# Batch Comparison Report\n\n"
        summary = []
        
        for i, (docx_file, url) in enumerate(matches, start=1):
            try:
                # Use the full path for processing
                draft_text = normalize_text(get_docx_text(docx_file))
                live_text, title, meta_desc = get_webpage_text(url)
                live_text = normalize_text(live_text)
                
                if "[ERROR" in live_text:
                    report_md += f"## {os.path.basename(docx_file)} vs {url}\n❌ {live_text}\n\n"
                    summary.append(f"❌ {url}: Error")
                    continue
                
                diff, similarity = block_compare(draft_text, live_text)
                
                # Generate reports using basename for display
                html_report = format_result_as_html(os.path.basename(docx_file), url, title, meta_desc, similarity, diff)
                markdown_report = format_result_as_markdown(os.path.basename(docx_file), url, title, meta_desc, similarity, diff)

                # Save HTML report
                html_file_path = os.path.join(results_folder, f"report_{i}_{os.path.splitext(os.path.basename(docx_file))[0]}.html")
                with open(html_file_path, "w", encoding="utf-8") as f:
                    f.write(f"""<!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset='UTF-8'>
                        <title>Verbatim AI - Comparison Report</title>
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

                report_md += markdown_report
                summary.append(f"{url} → Similarity: {similarity:.2%}")

            except Exception as e:
                report_md += f"## {os.path.basename(docx_file)} vs {url}\n❌ Error: {str(e)}\n\n"
                summary.append(f"❌ {url}: Error")
            
            progress_bar["value"] = i
            root.update_idletasks()
        
        # Save markdown report
        md_path = os.path.join(results_folder, "comparison_report.md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(report_md)
        
        # Store current report for saving later
        root.current_report_md = report_md
        
        # Update GUI
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, "Reports saved.\n\n" + "\n".join(summary))
        
        # Show completion message with theme-aware styling
        completion_window = tk.Toplevel(root)
        completion_window.title("Comparison Complete")
        
        # Make window modal
        completion_window.transient(root)
        completion_window.grab_set()
        
        # Set size and position
        window_width = 500
        window_height = 250  # Increased height to accommodate buttons
        screen_width = completion_window.winfo_screenwidth()
        screen_height = completion_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        completion_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Check if dark mode is enabled
        current_settings = load_settings()
        is_dark_mode = current_settings.get('dark_mode', 'false').lower() == 'true'
        
        # Apply theme colors
        bg_color = '#1e1e1e' if is_dark_mode else '#f0f0f0'
        fg_color = '#ffffff' if is_dark_mode else '#000000'
        
        # Configure window colors
        completion_window.configure(bg=bg_color)
        
        # Configure ttk styles
        style = ttk.Style()
        if is_dark_mode:
            style.configure('Complete.TFrame', background='#1e1e1e')
            style.configure('Complete.TButton',
                          background='#404040',
                          foreground='#000000',  # Changed to black
                          bordercolor='#505050',
                          lightcolor='#404040',
                          darkcolor='#2d2d2d',
                          relief='raised',
                          padding=(20, 10))
            
            style.map('Complete.TButton',
                     background=[('active', '#505050'), ('pressed', '#303030')],
                     foreground=[('active', '#000000'), ('pressed', '#000000')])  # Keep black on hover/press
        else:
            style.configure('Complete.TFrame', background='#f0f0f0')
            style.configure('Complete.TButton',
                          background='#e0e0e0',
                          foreground='#000000',
                          padding=(20, 10))
            
            style.map('Complete.TButton',
                     background=[('active', '#d0d0d0'), ('pressed', '#c0c0c0')],
                     foreground=[('active', '#000000'), ('pressed', '#000000')])

        # Create main frame
        main_frame = ttk.Frame(completion_window, style='Complete.TFrame')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)

        # Add checkmark and title
        title_text = "✅ Batch comparison complete"
        title_label = tk.Label(
            main_frame,
            text=title_text,
            font=("Roboto", 16, "bold"),
            bg=bg_color,
            fg=fg_color
        )
        title_label.pack(pady=(0, 20))

        # Add file paths
        path_text = f"Markdown saved to:\n{md_path}\n\nHTML reports saved in the same folder."
        path_label = tk.Label(
            main_frame,
            text=path_text,
            justify=tk.LEFT,
            bg=bg_color,
            fg=fg_color,
            wraplength=450
        )
        path_label.pack(pady=(0, 30))

        # Add buttons frame
        button_frame = ttk.Frame(main_frame, style='Complete.TFrame')
        button_frame.pack(fill='x', pady=(10, 0))

        def open_folder():
            os.startfile(results_folder)

        # Create and pack buttons with equal width
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        open_button = ttk.Button(
            button_frame,
            text="Open Results Folder",
            command=open_folder,
            style='Complete.TButton'
        )
        open_button.grid(row=0, column=0, padx=10, sticky='ew')

        close_button = ttk.Button(
            button_frame,
            text="Close",
            command=completion_window.destroy,
            style='Complete.TButton'
        )
        close_button.grid(row=0, column=1, padx=10, sticky='ew')

        # Center the window
        completion_window.update_idletasks()
        completion_window.geometry(f"+{x}+{y}")
        
        # Bind Enter and Escape keys
        completion_window.bind('<Return>', lambda e: open_folder())
        completion_window.bind('<Escape>', lambda e: completion_window.destroy())
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during comparison: {str(e)}")
        progress_bar["value"] = 0
    
    finally:
        # Re-enable the drop target
        if USE_DND:
            drop_target.drop_target_register(tkdnd.DND_FILES)
            drop_target.dnd_bind('<<Drop>>', handle_drop)

# ------------------ GUI Setup ------------------

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_last_report_location():
    """Get the most recent report location"""
    if hasattr(root, 'last_report_location') and os.path.exists(root.last_report_location):
        return root.last_report_location
    settings = load_settings()
    if settings['default_save_location'] and os.path.exists(settings['default_save_location']):
        return settings['default_save_location']
    return os.getcwd()

def open_report():
    """Open an existing HTML or Markdown report"""
    initial_dir = get_last_report_location()
    
    filetypes = [
        ("Report files", "*.html;*.md"),
        ("HTML files", "*.html"),
        ("Markdown files", "*.md"),
        ("All files", "*.*")
    ]
    
    filename = filedialog.askopenfilename(
        title="Open Report",
        filetypes=filetypes,
        initialdir=initial_dir
    )
    if filename:
        # Store the directory for next time
        root.last_report_location = os.path.dirname(filename)
        webbrowser.open(filename)

def save_report():
    """Save the current comparison results"""
    if not hasattr(root, 'current_report_md'):
        messagebox.showinfo("No Report", "No comparison results to save. Please run a comparison first.")
        return
    
    filetypes = [
        ("Markdown files", "*.md"),
        ("HTML files", "*.html"),
        ("All files", "*.*")
    ]
    
    initial_dir = get_last_report_location()
    
    filename = filedialog.asksaveasfilename(
        title="Save Report",
        filetypes=filetypes,
        defaultextension=".md",
        initialdir=initial_dir
    )
    if filename:
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(root.current_report_md)
            # Store the directory for next time
            root.last_report_location = os.path.dirname(filename)
            messagebox.showinfo("Success", "Report saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save report: {str(e)}")

def save_settings(settings):
    """Save settings to a file"""
    settings_dir = os.path.join(os.path.dirname(__file__), "config")
    if not os.path.exists(settings_dir):
        os.makedirs(settings_dir)
    
    settings_file = os.path.join(settings_dir, "settings.txt")
    with open(settings_file, 'w') as f:
        for key, value in settings.items():
            f.write(f"{key}={value}\n")

def load_settings():
    """Load settings from file"""
    # Get default Downloads folder path
    downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    
    settings = {
        'default_save_location': downloads_path,
        'similarity_threshold': '0.9',
        'dark_mode': 'false'
    }
    
    settings_file = os.path.join(os.path.dirname(__file__), "config", "settings.txt")
    if os.path.exists(settings_file):
        with open(settings_file, 'r') as f:
            for line in f:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    # Only update if the value is not empty and the path exists (for save location)
                    if value and (key != 'default_save_location' or os.path.exists(value)):
                        settings[key] = value
    
    return settings

def apply_theme(is_dark_mode):
    """Apply light or dark theme to the application"""
    if is_dark_mode:
        # Dark mode colors
        root.configure(bg='#1e1e1e')
        frame.configure(bg='#1e1e1e')
        style = ttk.Style()
        style.configure('TFrame', background='#1e1e1e')
        style.configure('TLabel', background='#1e1e1e', foreground='#ffffff')
        
        # Configure all button styles with black text in dark mode
        style.configure('TButton',
                      background='#404040',
                      foreground='#000000')
        style.map('TButton',
                 background=[('active', '#505050'), ('pressed', '#303030')],
                 foreground=[('active', '#000000'), ('pressed', '#000000')])
        
        style.configure('Select.TButton',
                      background='#404040',
                      foreground='#000000',
                      bordercolor='#505050',
                      lightcolor='#404040')
        style.map('Select.TButton',
                 background=[('active', '#505050'), ('pressed', '#303030')],
                 foreground=[('active', '#000000'), ('pressed', '#000000')])
        
        style.configure('url.TButton',
                      background='#404040',
                      foreground='#000000',
                      bordercolor='#505050',
                      lightcolor='#404040')
        style.map('url.TButton',
                 background=[('active', '#505050'), ('pressed', '#303030')],
                 foreground=[('active', '#000000'), ('pressed', '#000000')])
        
        style.configure('Complete.TButton',
                      background='#404040',
                      foreground='#000000',
                      bordercolor='#505050',
                      lightcolor='#404040')
        style.map('Complete.TButton',
                 background=[('active', '#505050'), ('pressed', '#303030')],
                 foreground=[('active', '#000000'), ('pressed', '#000000')])
        
        style.configure('TEntry', fieldbackground='#2d2d2d', foreground='#ffffff')
        
        # Configure menu colors
        menubar.configure(bg='#2d2d2d', fg='#ffffff', 
                        activebackground='#404040', activeforeground='#ffffff')
        for menu in [file_menu, edit_menu, view_menu, help_menu]:
            menu.configure(bg='#2d2d2d', fg='#ffffff', 
                         activebackground='#404040', activeforeground='#ffffff',
                         selectcolor='#ffffff')
        
        # Configure text area
        text_area.configure(
            bg='#2d2d2d',
            fg='#ffffff',
            insertbackground='#ffffff',
            selectbackground='#404040',
            selectforeground='#ffffff'
        )
        
        # Configure other widgets
        if hasattr(root, 'title_label'):
            root.title_label.configure(bg='#1e1e1e', fg='#ffffff')
        
        if USE_DND and hasattr(root, 'drop_target'):
            root.drop_target.configure(bg='#2d2d2d', fg='#ffffff')
        elif hasattr(root, 'no_dnd_label'):
            root.no_dnd_label.configure(bg='#2d2d2d', fg='#ffffff')

        # Update logo background for dark mode
        if hasattr(root, 'logo_label'):
            root.logo_label.configure(bg='#1e1e1e')
        
        # Update all themed entry widgets
        style.configure('url.TEntry', 
                       fieldbackground='#2d2d2d',
                       foreground='#000000')  # Changed to black
        
        # Update button text colors
        style.configure('TButton',
                       foreground='#000000')  # Changed to black
        style.configure('url.TButton',
                       foreground='#000000')  # Changed to black
        style.configure('Complete.TButton',
                       foreground='#000000')  # Changed to black
        
        # Update button hover states
        style.map('TButton',
                 foreground=[('active', '#000000'), ('pressed', '#000000')])
        style.map('url.TButton',
                 foreground=[('active', '#000000'), ('pressed', '#000000')])
        style.map('Complete.TButton',
                 foreground=[('active', '#000000'), ('pressed', '#000000')])
    else:
        # Light mode colors
        root.configure(bg='#f0f0f0')
        frame.configure(bg='#f0f0f0')
        style = ttk.Style()
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', foreground='#000000')
        
        style.configure('TButton',
                       background='#ffffff',
                       foreground='#000000')
        
        style.configure('Select.TButton',
                       background='#f0f0f0',
                       foreground='#000000')
        
        style.configure('url.TButton',
                       background='#f0f0f0',
                       foreground='#000000')
        
        style.configure('Complete.TButton',
                       background='#f0f0f0',
                       foreground='#000000')
        
        style.configure('TEntry', fieldbackground='#ffffff', foreground='#000000')
        
        # Configure menu colors
        menubar.configure(bg='#f0f0f0', fg='#000000', 
                        activebackground='#e0e0e0', activeforeground='#000000')
        for menu in [file_menu, edit_menu, view_menu, help_menu]:
            menu.configure(bg='#f0f0f0', fg='#000000', 
                         activebackground='#e0e0e0', activeforeground='#000000',
                         selectcolor='#000000')
        
        # Configure text area
        text_area.configure(
            bg='#ffffff',
            fg='#000000',
            insertbackground='#000000',
            selectbackground='#0078d7',
            selectforeground='#ffffff'
        )
        
        # Configure other widgets
        if hasattr(root, 'title_label'):
            root.title_label.configure(bg='#f0f0f0', fg='#000000')
        
        if USE_DND and hasattr(root, 'drop_target'):
            root.drop_target.configure(bg='#f0f0f0', fg='#000000')
        elif hasattr(root, 'no_dnd_label'):
            root.no_dnd_label.configure(bg='#f0f0f0', fg='#000000')

        # Update logo background for light mode
        if hasattr(root, 'logo_label'):
            root.logo_label.configure(bg='#f0f0f0')

def toggle_dark_mode():
    """Toggle between light and dark mode"""
    current_settings = load_settings()
    is_dark_mode = current_settings.get('dark_mode', 'false').lower() == 'true'
    
    # Toggle the mode
    is_dark_mode = not is_dark_mode
    
    # Apply the theme
    apply_theme(is_dark_mode)
    
    # Save the setting
    current_settings['dark_mode'] = str(is_dark_mode).lower()
    save_settings(current_settings)

def show_settings():
    """Show settings/preferences dialog"""
    settings_window = tk.Toplevel(root)
    settings_window.title("Settings")
    settings_window.geometry("500x400")
    
    # Make window modal
    settings_window.transient(root)
    settings_window.grab_set()
    
    # Center the window
    window_width = 500
    window_height = 400
    screen_width = settings_window.winfo_screenwidth()
    screen_height = settings_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    settings_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Create notebook for tabbed interface
    notebook = ttk.Notebook(settings_window)
    notebook.pack(fill='both', expand=True, padx=10, pady=10)
    
    # General Settings tab
    general_frame = ttk.Frame(notebook)
    notebook.add(general_frame, text='General')
    
    # Comparison Settings tab
    comparison_frame = ttk.Frame(notebook)
    notebook.add(comparison_frame, text='Comparison')
    
    # Load current settings
    current_settings = load_settings()
    
    # Add settings controls
    ttk.Label(general_frame, text="Default Save Location:").pack(anchor='w', padx=10, pady=5)
    save_location = ttk.Entry(general_frame)
    save_location.pack(fill='x', padx=10)
    if current_settings['default_save_location']:
        save_location.insert(0, current_settings['default_save_location'])
    
    def browse_save_location():
        folder = filedialog.askdirectory(title="Select Default Save Location")
        if folder:
            save_location.delete(0, tk.END)
            save_location.insert(0, folder)
    
    ttk.Button(general_frame, text="Browse...", command=browse_save_location).pack(anchor='w', padx=10)
    
    ttk.Label(comparison_frame, text="Similarity Threshold:").pack(anchor='w', padx=10, pady=5)
    similarity_scale = ttk.Scale(comparison_frame, from_=0.5, to=1.0, orient='horizontal')
    similarity_scale.set(float(current_settings['similarity_threshold']))
    similarity_scale.pack(fill='x', padx=10)
    
    def save_and_close():
        # Save settings
        new_settings = {
            'default_save_location': save_location.get(),
            'similarity_threshold': str(similarity_scale.get())
        }
        save_settings(new_settings)
        settings_window.destroy()
    
    def cancel():
        settings_window.destroy()
    
    # Add buttons
    button_frame = ttk.Frame(settings_window)
    button_frame.pack(fill='x', padx=10, pady=10)
    ttk.Button(button_frame, text="Save", command=save_and_close).pack(side='right', padx=5)
    ttk.Button(button_frame, text="Cancel", command=cancel).pack(side='right', padx=5)

def copy_results():
    """Copy the current results to clipboard"""
    if not text_area.get(1.0, tk.END).strip():
        messagebox.showinfo("No Results", "No results to copy. Please run a comparison first.")
        return
    root.clipboard_clear()
    root.clipboard_append(text_area.get(1.0, tk.END))
    root.update()

def clear_all():
    """Clear all results and reset the interface"""
    if text_area.get(1.0, tk.END).strip():
        if messagebox.askyesno("Clear All", "Are you sure you want to clear all results?"):
            text_area.delete(1.0, tk.END)
            progress_bar["value"] = 0
            if hasattr(root, 'current_report_md'):
                delattr(root, 'current_report_md')

def show_documentation():
    """Show the documentation in the default web browser"""
    docs_dir = os.path.join(os.path.dirname(__file__), "docs")
    docs_path = os.path.join(docs_dir, "html_report_guide.md")
    
    # Create docs directory if it doesn't exist
    if not os.path.exists(docs_dir):
        os.makedirs(docs_dir)
    
    # Create documentation file if it doesn't exist
    if not os.path.exists(docs_path):
        with open(docs_path, 'w', encoding='utf-8') as f:
            f.write("""# Verbatim AI Documentation

## Overview
Verbatim AI is a powerful tool for comparing draft content with live website content. This guide will help you understand how to use the application effectively.

## Getting Started
1. Launch Verbatim AI
2. Click "Upload Documents" or drag and drop DOCX files
3. Enter the corresponding URLs for each document
4. Click "Start AutoCompare" to begin the comparison

## Understanding the Reports
- Green: Content matches between draft and live site
- Red: Content in draft but missing from live site
- Blue: Content on live site but not in draft

## Features
- Batch processing of multiple documents
- Side-by-side comparison view
- Similarity scoring
- Export options (HTML and Markdown)
- Keyboard shortcuts

## Keyboard Shortcuts
- Ctrl+N: New Comparison
- Ctrl+O: Open Report
- Ctrl+S: Save Report
- Ctrl+C: Copy Results
- Ctrl+Delete: Clear All
- F1: Open Documentation
- Alt+F4: Exit

## Support
For additional support or to report issues, please contact the SMB Team.

© 2025 SMB Team. All rights reserved.""")
    
    webbrowser.open(f"file://{docs_path}")

def show_about():
    """Show the About dialog"""
    about_window = tk.Toplevel(root)
    about_window.title("About Verbatim AI")
    about_window.geometry("400x300")
    
    # Make window modal
    about_window.transient(root)
    about_window.grab_set()
    
    # Center the window
    window_width = 400
    window_height = 300
    screen_width = about_window.winfo_screenwidth()
    screen_height = about_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    about_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Add logo if available
    try:
        logo_path = resource_path('smbteam-logo.png')
        logo_image = tk.PhotoImage(file=logo_path)
        # Resize the image to a reasonable size
        logo_image = logo_image.subsample(3, 3)
        logo_label = tk.Label(about_window, image=logo_image)
        logo_label.image = logo_image
        logo_label.pack(pady=10)
    except Exception:
        pass
    
    # Add version and copyright information
    tk.Label(about_window, text="Verbatim AI", font=("Roboto", 16, "bold")).pack(pady=5)
    tk.Label(about_window, text="Version 1.0.0").pack()
    tk.Label(about_window, text="© 2025 SMB Team. All rights reserved.").pack(pady=5)
    tk.Label(about_window, text="A powerful tool for comparing draft content\nwith live website content.", justify=tk.CENTER).pack(pady=10)
    
    # Add close button
    ttk.Button(about_window, text="Close", command=about_window.destroy).pack(pady=10)

def crawl_website(base_url, max_pages=100):
    """Crawl a website starting from base_url and extract URLs with their content"""
    import urllib.parse
    from urllib.robotparser import RobotFileParser
    import time
    from concurrent.futures import ThreadPoolExecutor, as_completed
    
    # Normalize base URL
    base_url = base_url.rstrip('/')
    base_domain = urllib.parse.urlparse(base_url).netloc
    
    # Initialize robot parser
    robots_url = urllib.parse.urljoin(base_url, '/robots.txt')
    rp = RobotFileParser()
    rp.set_url(robots_url)
    try:
        rp.read()
    except:
        pass  # If robots.txt is not accessible, we'll proceed with crawling
    
    # Initialize variables
    visited_urls = set()
    urls_to_visit = {base_url}
    page_contents = {}
    
    # Add delay between requests
    request_delay = 0.5  # seconds
    
    # Headers for requests
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Connection': 'keep-alive',
    }

    def normalize_url(url):
        """Normalize URL to avoid duplicates"""
        parsed = urllib.parse.urlparse(url)
        # Remove trailing slashes and default ports
        normalized = parsed.scheme + '://' + parsed.netloc.rstrip(':80').rstrip(':443') + parsed.path.rstrip('/')
        # Add query parameters if they exist
        if parsed.query:
            normalized += '?' + parsed.query
        return normalized

    def extract_links(soup, current_url):
        """Extract all possible links from the page"""
        links = set()
        
        # Find all elements that might contain links
        for element in soup.find_all(['a', 'link', 'area', 'base', 'nav', 'menu']):
            # Check href attribute
            href = element.get('href')
            if href:
                try:
                    absolute_url = urllib.parse.urljoin(current_url, href)
                    parsed_url = urllib.parse.urlparse(absolute_url)
                    
                    # Only include URLs from the same domain and with http(s) scheme
                    if (parsed_url.netloc == base_domain and 
                        parsed_url.scheme in ('http', 'https') and
                        not any(ext in parsed_url.path.lower() for ext in ['.jpg', '.jpeg', '.png', '.gif', '.pdf', '.doc', '.docx'])):
                        
                        normalized_url = normalize_url(absolute_url)
                        links.add(normalized_url)
                except:
                    continue
        
        # Also look for links in navigation menus and other structures
        for menu_item in soup.find_all(class_=lambda x: x and any(word in str(x).lower() for word in ['menu', 'nav', 'navigation'])):
            for link in menu_item.find_all('a', href=True):
                try:
                    absolute_url = urllib.parse.urljoin(current_url, link['href'])
                    if urllib.parse.urlparse(absolute_url).netloc == base_domain:
                        normalized_url = normalize_url(absolute_url)
                        links.add(normalized_url)
                except:
                    continue
        
        return links
    
    def fetch_url(url):
        try:
            time.sleep(request_delay)
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            
            # Check if it's an HTML page
            content_type = response.headers.get('content-type', '').lower()
            if 'text/html' not in content_type:
                return None
            
            content, title, meta_desc = get_webpage_text(url)
            if "[ERROR" not in content:
                # Extract links from the page
                soup = BeautifulSoup(response.text, 'html.parser')
                links = extract_links(soup, url)
                
                return url, {
                    'content': content,
                    'title': title,
                    'meta_desc': meta_desc,
                    'links': links
                }
        except Exception as e:
            print(f"Error fetching {url}: {str(e)}")
        return None
    
    # Use ThreadPoolExecutor for parallel crawling
    with ThreadPoolExecutor(max_workers=5) as executor:
        while urls_to_visit and len(visited_urls) < max_pages:
            # Get next batch of URLs to process
            batch = set()
            while urls_to_visit and len(batch) < 5:
                url = urls_to_visit.pop()
                normalized_url = normalize_url(url)
                if normalized_url not in visited_urls and rp.can_fetch('*', url):
                    batch.add(url)
                    visited_urls.add(normalized_url)
            
            if not batch:
                break
            
            # Process batch in parallel
            future_to_url = {executor.submit(fetch_url, url): url for url in batch}
            
            for future in as_completed(future_to_url):
                result = future.result()
                if result:
                    url, data = result
                    normalized_url = normalize_url(url)
                    page_contents[normalized_url] = data
                    # Add new URLs to visit
                    for new_url in data['links']:
                        if normalize_url(new_url) not in visited_urls:
                            urls_to_visit.add(new_url)
    
    return page_contents

def handle_unmatched_document(parent_window, docx_file, potential_matches=None):
    """Handle a document that couldn't be automatically matched"""
    dialog = tk.Toplevel(parent_window)
    dialog.title("Document Match Not Found")
    dialog.geometry("600x500")  # Increased height from 400 to 500
    
    # Make dialog modal
    dialog.transient(parent_window)
    dialog.grab_set()
    
    # Center the dialog
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f'+{x}+{y}')
    
    # Check if dark mode is enabled
    current_settings = load_settings()
    is_dark_mode = current_settings.get('dark_mode', 'false').lower() == 'true'
    
    # Configure styles for dark mode
    style = ttk.Style()
    if is_dark_mode:
        dialog.configure(bg='#1e1e1e')
        style.configure('Dialog.TFrame', background='#1e1e1e')
        style.configure('Dialog.TLabel', background='#1e1e1e', foreground='#ffffff')
        style.configure('Dialog.TLabelframe', background='#1e1e1e', foreground='#ffffff')
        style.configure('Dialog.TLabelframe.Label', background='#1e1e1e', foreground='#ffffff')
        style.configure('Dialog.TRadiobutton', background='#1e1e1e', foreground='#ffffff')
        style.configure('Dialog.TEntry', 
                       fieldbackground='#2d2d2d',
                       foreground='#000000',  # Keep text black for readability
                       insertcolor='#ffffff')
        style.configure('Dialog.TButton',
                       background='#404040',
                       foreground='#000000')  # Keep button text black
    else:
        style.configure('Dialog.TFrame', background='#f0f0f0')
        style.configure('Dialog.TLabel', background='#f0f0f0', foreground='#000000')
        style.configure('Dialog.TLabelframe', background='#f0f0f0', foreground='#000000')
        style.configure('Dialog.TLabelframe.Label', background='#f0f0f0', foreground='#000000')
        style.configure('Dialog.TRadiobutton', background='#f0f0f0', foreground='#000000')
        style.configure('Dialog.TEntry', 
                       fieldbackground='#ffffff',
                       foreground='#000000',
                       insertcolor='#000000')
        style.configure('Dialog.TButton',
                       background='#e0e0e0',
                       foreground='#000000')
    
    # Configure grid
    dialog.grid_columnconfigure(0, weight=1)
    
    # Add warning icon and message
    frame = ttk.Frame(dialog, style='Dialog.TFrame')
    frame.grid(row=0, column=0, padx=20, pady=10, sticky='nsew')
    
    warning_label = ttk.Label(
        frame,
        text=f"⚠️ Could not automatically match:\n{os.path.basename(docx_file)}",
        font=("Roboto", 11, "bold"),
        wraplength=500,
        style='Dialog.TLabel'
    )
    warning_label.pack(pady=(0, 10))
    
    # Add potential matches section if available
    result = {"action": None, "url": None}
    
    if potential_matches:
        matches_frame = ttk.LabelFrame(frame, text="Potential Matches", padding=10, style='Dialog.TLabelframe')
        matches_frame.pack(fill='x', pady=10)
        
        # Create scrollable frame for matches
        canvas = tk.Canvas(matches_frame, height=150, 
                         bg='#1e1e1e' if is_dark_mode else '#f0f0f0',
                         highlightthickness=0)
        scrollbar = ttk.Scrollbar(matches_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='Dialog.TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=500)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Add matches with radio buttons
        selected_match = tk.StringVar()
        for url, similarity in potential_matches:
            match_frame = ttk.Frame(scrollable_frame, style='Dialog.TFrame')
            match_frame.pack(fill='x', pady=2)
            
            rb = ttk.Radiobutton(
                match_frame,
                text=f"{url}\nSimilarity: {similarity:.2%}",
                value=url,
                variable=selected_match,
                style='Dialog.TRadiobutton'
            )
            rb.pack(fill='x')
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create a separate frame for the Use Selected Match button
        button_frame = ttk.Frame(frame, style='Dialog.TFrame')
        button_frame.pack(fill='x', pady=5)
        
        def use_selected_match():
            if selected_match.get():
                result["action"] = "use_match"
                result["url"] = selected_match.get()
                dialog.destroy()
        
        ttk.Button(
            button_frame,
            text="Use Selected Match",
            command=use_selected_match,
            style='Dialog.TButton'
        ).pack(side='right', padx=5)
    
    # Add manual URL entry section
    manual_frame = ttk.LabelFrame(frame, text="Enter URL Manually", padding=10, style='Dialog.TLabelframe')
    manual_frame.pack(fill='x', pady=10)
    
    url_entry = ttk.Entry(manual_frame, width=50, style='Dialog.TEntry')
    url_entry.pack(side='left', padx=5)
    
    def use_manual_url():
        if url_entry.get().strip():
            result["action"] = "manual"
            result["url"] = url_entry.get().strip()
            dialog.destroy()
    
    ttk.Button(
        manual_frame,
        text="Use This URL",
        command=use_manual_url,
        style='Dialog.TButton'
    ).pack(side='left', padx=5)
    
    # Add skip button at the bottom
    skip_frame = ttk.Frame(frame, style='Dialog.TFrame')
    skip_frame.pack(fill='x', pady=10)
    
    def skip_document():
        result["action"] = "skip"
        dialog.destroy()
    
    ttk.Button(
        skip_frame,
        text="Skip This Document",
        command=skip_document,
        style='Dialog.TButton'
    ).pack(side='right', padx=5)
    
    # Wait for dialog to close
    dialog.wait_window()
    return result["action"], result["url"]

def get_content_similarity(text1, text2):
    """Calculate content similarity between two texts using sequence matcher"""
    # Normalize both texts
    text1 = normalize_text(text1)
    text2 = normalize_text(text2)
    
    # Use SequenceMatcher for similarity
    return difflib.SequenceMatcher(None, text1, text2).ratio()

def get_url_path_similarity(url, docx_path):
    """Calculate similarity between URL path and document name"""
    # Extract the path part from URL
    url_path = urllib.parse.urlparse(url).path.strip('/')
    
    # Get document name without extension
    doc_name = os.path.splitext(os.path.basename(docx_path))[0]
    
    # Convert both to lowercase and replace special characters with spaces
    url_path = re.sub(r'[-_/]', ' ', url_path.lower())
    doc_name = re.sub(r'[-_]', ' ', doc_name.lower())
    
    # Calculate similarity
    return difflib.SequenceMatcher(None, url_path, doc_name).ratio()

def toggle_theme(mode):
    """Toggle between light and dark mode"""
    current_settings = load_settings()
    is_dark_mode = mode == "dark"
    
    # Apply the theme
    apply_theme(is_dark_mode)
    
    # Update menu checkmarks
    view_menu.entryconfigure("Light Mode", state="normal" if is_dark_mode else "disabled")
    view_menu.entryconfigure("Dark Mode", state="disabled" if is_dark_mode else "normal")
    
    # Save the setting
    current_settings['dark_mode'] = str(is_dark_mode).lower()
    save_settings(current_settings)

def auto_match_documents(docx_files, base_url, parent_window):
    """Automatically match DOCX files to URLs based on content similarity"""
    # Show progress window
    progress_window = tk.Toplevel(parent_window)
    progress_window.title("Matching Documents to URLs")
    progress_window.geometry("500x200")
    
    # Check if dark mode is enabled
    current_settings = load_settings()
    is_dark_mode = current_settings.get('dark_mode', 'false').lower() == 'true'
    
    # Apply theme colors
    bg_color = '#1e1e1e' if is_dark_mode else '#f0f0f0'
    fg_color = '#ffffff' if is_dark_mode else '#000000'
    
    # Configure window colors
    progress_window.configure(bg=bg_color)
    
    # Configure styles
    style = ttk.Style()
    
    # Create a custom style for the progress bar
    style.configure('Custom.Horizontal.TProgressbar',
                   troughcolor='#2d2d2d' if is_dark_mode else '#e0e0e0',
                   background='#404040' if is_dark_mode else '#0078d7',
                   darkcolor='#404040' if is_dark_mode else '#0078d7',
                   lightcolor='#404040' if is_dark_mode else '#0078d7',
                   bordercolor='#2d2d2d' if is_dark_mode else '#e0e0e0')
    
    # Center the window
    window_width = 500
    window_height = 200
    screen_width = progress_window.winfo_screenwidth()
    screen_height = progress_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    progress_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Make window modal
    progress_window.transient(parent_window)
    progress_window.grab_set()
    
    # Add progress label with theme-aware styling
    progress_label = tk.Label(
        progress_window,
        text="Crawling website and matching documents...\nThis may take a few minutes.",
        bg=bg_color,
        fg=fg_color
    )
    progress_label.pack(pady=20)
    
    # Add status label with theme-aware styling
    status_label = tk.Label(
        progress_window,
        text="",
        bg=bg_color,
        fg=fg_color
    )
    status_label.pack(pady=5)
    
    # Add progress bar with theme-aware styling
    progress_bar = ttk.Progressbar(
        progress_window,
        orient="horizontal",
        length=400,
        mode="indeterminate",
        style='Custom.Horizontal.TProgressbar'
    )
    progress_bar.pack(pady=10)
    progress_bar.start()
    
    try:
        # Normalize base URL
        if not base_url.startswith(('http://', 'https://')):
            base_url = 'https://' + base_url
        base_url = base_url.rstrip('/')
        
        # Update status
        status_label.config(text="Crawling website...")
        progress_window.update()
        
        # Crawl the website
        page_contents = crawl_website(base_url)
        
        if not page_contents:
            raise Exception("No pages found to crawl. Please check the URL and try again.")
        
        # Update status
        status_label.config(text=f"Found {len(page_contents)} pages. Matching documents...")
        progress_window.update()
        
        # Process each DOCX file
        matches = []
        for docx_file in docx_files:
            try:
                # Extract text from DOCX
                docx_text = normalize_text(get_docx_text(docx_file))
                
                # Find potential matches
                potential_matches = []
                for url, content in page_contents.items():
                    content_sim = get_content_similarity(docx_text, content['content'])
                    url_sim = get_url_path_similarity(url, docx_file)
                    
                    # Calculate combined similarity
                    if url_sim > 0.8:  # High URL match
                        combined_sim = url_sim * 0.7 + content_sim * 0.3
                    else:
                        combined_sim = url_sim * 0.3 + content_sim * 0.7
                    
                    potential_matches.append((url, combined_sim))
                
                # Sort matches by similarity
                potential_matches.sort(key=lambda x: x[1], reverse=True)
                
                # Check if we have a confident match
                if potential_matches and potential_matches[0][1] > 0.3:
                    matches.append((docx_file, potential_matches[0][0]))
                    status_label.config(text=f"Matched: {os.path.basename(docx_file)}")
                else:
                    # Hide progress window temporarily
                    progress_window.withdraw()
                    
                    # Show unmatched document dialog
                    action, url = handle_unmatched_document(
                        parent_window,
                        docx_file,
                        potential_matches[:5] if potential_matches else None  # Show top 5 matches
                    )
                    
                    # Show progress window again
                    progress_window.deiconify()
                    
                    if action == "use_match" or action == "manual":
                        matches.append((docx_file, url))
                        status_label.config(text=f"Matched: {os.path.basename(docx_file)} (Manual/Selected)")
                    elif action == "skip":
                        status_label.config(text=f"Skipped: {os.path.basename(docx_file)}")
                
                progress_window.update()
                
            except Exception as e:
                progress_window.destroy()
                messagebox.showerror("Error", f"Error processing {os.path.basename(docx_file)}: {str(e)}")
                return None
        
        if not matches:
            progress_window.destroy()
            messagebox.showerror("Error", "No matches found for any documents.")
            return None
        
        progress_window.destroy()
        return matches
        
    except Exception as e:
        progress_window.destroy()
        messagebox.showerror("Error", f"Failed to match documents: {str(e)}")
        return None

if __name__ == "__main__":
    # Use tkdnd.Tk if available, otherwise use regular tk.Tk
    root = tkdnd.Tk() if USE_DND else tk.Tk()
    root.title("Verbatim AI")

    # Create menu bar
    menubar = tk.Menu(root)
    root.config(menu=menubar)

    # Create File menu
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    
    # Add File menu items
    file_menu.add_command(label="New Comparison", command=run_batch_comparison, accelerator="Ctrl+N")
    file_menu.add_separator()
    file_menu.add_command(label="Settings", command=show_settings, accelerator="Ctrl+,")
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit, accelerator="Alt+F4")

    # Create Edit menu
    edit_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Edit", menu=edit_menu)
    
    # Add Edit menu items
    edit_menu.add_command(label="Copy Results", command=copy_results, accelerator="Ctrl+C")
    edit_menu.add_separator()
    edit_menu.add_command(label="Clear All", command=clear_all, accelerator="Ctrl+Delete")

    # Create View menu
    view_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="View", menu=view_menu)
    
    # Add View menu items - Light/Dark mode options
    current_settings = load_settings()
    is_dark_mode = current_settings.get('dark_mode', 'false').lower() == 'true'
    view_menu.add_command(
        label="Light Mode",
        command=lambda: toggle_theme("light"),
        state="normal" if is_dark_mode else "disabled"
    )
    view_menu.add_command(
        label="Dark Mode",
        command=lambda: toggle_theme("dark"),
        state="disabled" if is_dark_mode else "normal"
    )

    # Create Help menu
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    
    # Add Help menu items
    help_menu.add_command(label="Documentation", command=show_documentation, accelerator="F1")
    help_menu.add_separator()
    help_menu.add_command(label="About Verbatim AI", command=show_about)

    # Bind keyboard shortcuts
    root.bind("<Control-n>", lambda e: run_batch_comparison())
    root.bind("<Control-o>", lambda e: open_report())
    root.bind("<Control-s>", lambda e: save_report())
    root.bind("<Control-comma>", lambda e: show_settings())
    root.bind("<Control-c>", lambda e: copy_results())
    root.bind("<Control-Delete>", lambda e: clear_all())
    root.bind("<F1>", lambda e: show_documentation())

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
        logo_image = logo_image.subsample(2, 2)
        logo_label = tk.Label(frame, image=logo_image)
        logo_label.image = logo_image
        root.logo_label = logo_label  # Store reference for theme switching
        logo_label.pack(pady=(0, 10))
    except Exception as e:
        print(f"Error loading logo: {e}")
        pass

    # Add title text with Roboto font
    title_label = tk.Label(frame, text="Verbatim AI", font=("Roboto", 32, "bold"))
    root.title_label = title_label  # Store reference for theme switching
    title_label.pack(pady=(0, 20))

    # Create drop target area if DND is available
    if USE_DND:
        drop_target = tk.Label(
            frame,
            text="Drag and drop DOCX files or folders here",
            width=40,
            height=5,
            relief="solid",
            borderwidth=2
        )
        root.drop_target = drop_target  # Store reference for theme switching
        drop_target.pack(pady=10)
        
        # Register drop target
        drop_target.drop_target_register(tkdnd.DND_FILES)
        drop_target.dnd_bind('<<Drop>>', handle_drop)
    else:
        # If DND is not available, show a message
        no_dnd_label = tk.Label(
            frame,
            text="Drag and drop not available in this environment.\nPlease use the buttons below to select files.",
            width=40,
            height=5,
            relief="solid",
            borderwidth=2
        )
        no_dnd_label.pack(pady=10)

    # Create button frame
    button_frame = ttk.Frame(frame)
    button_frame.pack(pady=5)

    def select_folder_direct():
        folder_path = filedialog.askdirectory(title="Select Folder Containing Draft DOCX Files")
        if folder_path:
            process_files(folder_path)

    def select_files_direct():
        files = filedialog.askopenfilenames(
            title="Select DOCX Files",
            filetypes=[("DOCX files", "*.docx"), ("All files", "*.*")]
        )
        if files:
            # Create temporary folder for selected files
            temp_folder = os.path.join(os.path.dirname(files[0]), "VerbatimAI_Results")
            if not os.path.exists(temp_folder):
                os.makedirs(temp_folder)
            
            # Process selected files
            process_files(temp_folder, files)

    # Add the action buttons
    select_folder_btn = ttk.Button(button_frame, text="Select Folder", command=select_folder_direct)
    select_folder_btn.pack(side="left", padx=5)
    
    select_files_btn = ttk.Button(button_frame, text="Select Files", command=select_files_direct)
    select_files_btn.pack(side="left", padx=5)

    # Add progress bar
    progress_bar = ttk.Progressbar(frame, orient="horizontal", length=600, mode="determinate")
    progress_bar.pack(pady=5)

    # Add text area
    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=10)
    text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Apply initial theme based on settings
    current_settings = load_settings()
    is_dark_mode = current_settings.get('dark_mode', 'false').lower() == 'true'
    apply_theme(is_dark_mode)

    root.mainloop() 