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
    window_width = 1200
    window_height = 600
    match_window.geometry(f"{window_width}x{window_height}")
    entries = []
    canvas = tk.Canvas(match_window, width=window_width)
    scrollbar = tk.Scrollbar(match_window, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    tk.Label(scroll_frame, text="Enter the URL that matches each DOCX file:", font=("Arial", 12, "bold")).pack(pady=10)
    for file in docx_files:
        frame = tk.Frame(scroll_frame)
        frame.pack(fill="x", padx=10, pady=5)
        tk.Label(frame, text=file, width=40, anchor="w").pack(side="left")
        url_entry = tk.Entry(frame, width=100)
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
                # Preserve heading tags
                if tag.name.startswith('h'):
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
    # Only count content after H1 tags to exclude metadata
    total_draft_length = sum(len(block) for block in draft_blocks[draft_h1_index:]) if draft_h1_index != -1 else sum(len(block) for block in draft_blocks)
    total_live_length = sum(len(block) for block in live_blocks[live_h1_index:]) if live_h1_index != -1 else sum(len(block) for block in live_blocks)
    matched_content_length = 0
    
    # If we found H1s in both, align them
    if draft_h1_index != -1 and live_h1_index != -1:
        # Create aligned blocks with proper spacing
        aligned = []
        
        # Add draft blocks before H1 as unmatched
        for i in range(draft_h1_index):
            aligned.append(("missing", draft_blocks[i], ""))
        
        # Add live blocks before H1 as current
        for i in range(live_h1_index):
            aligned.append(("current", "", live_blocks[i]))
        
        # Now match the remaining blocks starting from H1
        draft_blocks = draft_blocks[draft_h1_index:]
        live_blocks = live_blocks[live_h1_index:]
    
    # First pass: try to match complete paragraphs
    matched_live = set()
    if 'aligned' not in locals():
        aligned = []
    
    for db in draft_blocks:
        best_match = None
        best_score = 0
        
        # Try to find the best matching block
        for lb in live_blocks:
            if lb in matched_live:
                continue
            score = difflib.SequenceMatcher(None, db, lb).ratio()
            if score > best_score:
                best_score = score
                best_match = lb
        
        # If we have a good match, use it
        if best_score >= similarity_threshold:
            matched_live.add(best_match)
            aligned.append(("matched", db, best_match))
            # Add the matched content length weighted by the match score
            matched_content_length += len(db) * best_score
        else:
            # If no good match, try to find partial matches
            partial_matches = []
            partial_match_length = 0
            for lb in live_blocks:
                if lb in matched_live:
                    continue
                # Split the live block into sentences
                live_sentences = [s.strip() for s in lb.split('.') if s.strip()]
                draft_sentences = [s.strip() for s in db.split('.') if s.strip()]
                
                # Check if any sentences match
                for ds in draft_sentences:
                    for ls in live_sentences:
                        match_score = difflib.SequenceMatcher(None, ds, ls).ratio()
                        if match_score > 0.8:  # Lower threshold for partial matches
                            partial_matches.append((ds, ls))
                            partial_match_length += len(ds) * match_score
            
            if partial_matches:
                # Combine partial matches into a single block
                combined_live = " ".join(m[1] for m in partial_matches)
                matched_live.add(combined_live)
                aligned.append(("matched", db, combined_live))
                matched_content_length += partial_match_length
            else:
                aligned.append(("missing", db, best_match if best_match else ""))

    # Add any unmatched live blocks
    for lb in live_blocks:
        if lb not in matched_live:
            aligned.append(("current", "", lb))

    # Calculate the weighted similarity score with more emphasis on matched content
    if total_draft_length == 0 or total_live_length == 0:
        similarity = 0.0
    else:
        # Calculate individual similarities
        draft_similarity = matched_content_length / total_draft_length
        live_similarity = matched_content_length / total_live_length
        
        # Use the higher similarity score to give more weight to matched content
        # Also consider the ratio of matched blocks to total blocks
        matched_blocks = sum(1 for tag, _, _ in aligned if tag == "matched")
        total_blocks = len(aligned)
        block_similarity = matched_blocks / total_blocks if total_blocks > 0 else 0
        
        # Combine the content similarity and block similarity
        similarity = max(draft_similarity, live_similarity) * 0.7 + block_similarity * 0.3
    
    # Return both the alignment results and the calculated similarity score
    return aligned, similarity

def format_result_as_html(docx_file, url, title, meta_desc, similarity, results):
    report = f"<h2>{docx_file} vs <a href='{url}'>{url}</a></h2>"
    report += f"<p><strong>Page Title:</strong> {title}</p>"
    report += f"<p><strong>Meta Description:</strong> {meta_desc}</p>"
    report += f"<p><strong>Similarity Score:</strong> {similarity:.2%}</p>"
    report += "<div style='display: flex; gap: 20px;'>"
    
    # Draft column
    report += "<div style='width: 50%;'><h3>Draft</h3><div style='white-space: pre-wrap;'>"
    for tag, draft, live in results:
        if tag in ['matched', 'missing']:
            report += f"<div style='background:{'#e6ffe6' if tag == 'matched' else '#ffe6e6'}; margin-bottom: 10px; padding: 10px;'>{draft}</div>"
        else:  # tag == 'current'
            # Create an invisible div with the same content to maintain spacing
            report += f"<div style='visibility: hidden; margin-bottom: 10px; padding: 10px;'>{live}</div>"
    report += "</div></div>"
    
    # Live column
    report += "<div style='width: 50%;'><h3>Live Webpage</h3><div style='white-space: pre-wrap;'>"
    for tag, draft, live in results:
        if tag in ['matched', 'current']:
            report += f"<div style='background:{'#e6ffe6' if tag == 'matched' else '#e6f0ff'}; margin-bottom: 10px; padding: 10px;'>{live}</div>"
        else:  # tag == 'missing'
            # Create an invisible div with the same content to maintain spacing
            report += f"<div style='visibility: hidden; margin-bottom: 10px; padding: 10px;'>{draft}</div>"
    report += "</div></div>"
    
    report += "</div><hr>"
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
                f.write(f"<html><head><meta charset='UTF-8'><title>Comparison Report</title></head><body>{html_report}</body></html>")
            if i == total:
                webbrowser.open(f"file://{html_file_path}")

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

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Manual Match Draft vs Webpage Comparison Tool")

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    button = tk.Button(frame, text="Run Manual Match & Compare", command=run_batch_comparison)
    button.pack(pady=5)

    progress_bar = ttk.Progressbar(frame, orient="horizontal", length=600, mode="determinate")
    progress_bar.pack(pady=5)

    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=10)
    text_area.pack(padx=10, pady=10)

    root.mainloop() 