import os
import sys
import argparse


def read_docx(file_path):
    from docx import Document
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text.append(cell.text)
        return '\n'.join(full_text)
    except Exception as e:
        print(f"Warning: Failed to read {file_path}: {str(e)}", file=sys.stderr)
        return ""


def read_doc(file_path, word_app):
    try:
        doc = word_app.Documents.Open(
            FileName=os.path.abspath(file_path),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False
        )
        text = doc.Content.Text
        doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
        return text
    except Exception as e:
        print(f"Warning: Failed to read {file_path}: {str(e)}", file=sys.stderr)
        return ""


def extract_text(file_path, word_app=None):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        return read_docx(file_path)
    elif ext == ".doc":
        return read_doc(file_path, word_app)
    return ""


def get_paragraphs_docx(file_path):
    from docx import Document
    try:
        doc = Document(file_path)
        paragraphs = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                paragraphs.append(text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    if text:
                        paragraphs.append(text)
        return paragraphs
    except Exception as e:
        print(f"Warning: Failed to read {file_path}: {str(e)}", file=sys.stderr)
        return []


def get_paragraphs_doc(file_path, word_app):
    try:
        doc = word_app.Documents.Open(
            FileName=os.path.abspath(file_path),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False
        )
        paragraphs = []
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if text:
                paragraphs.append(text)
        doc.Close(SaveChanges=0)
        return paragraphs
    except Exception as e:
        print(f"Warning: Failed to read {file_path}: {str(e)}", file=sys.stderr)
        return []


def is_english(text):
    english_chars = sum(1 for c in text if c.isalpha() and ord(c) < 128)
    total_chars = sum(1 for c in text if c.isalpha())
    if total_chars == 0:
        return False
    return english_chars / total_chars > 0.5


def truncate_text(text, max_length=300):
    if len(text) <= max_length:
        return text
    return text[:max_length] + '...'


def _build_translation(text, search_str, search_len, match_func, find_func, context_chars):
    """Build a translation excerpt, centering on the keyword if present."""
    if match_func(text, search_str):
        t_idx = find_func(text, search_str)
        t_match_end = t_idx + search_len
        t_start = max(0, t_idx - context_chars)
        t_end = min(len(text), t_match_end + context_chars)
        t_prefix = '...' if t_start > 0 else ''
        t_suffix = '...' if t_end < len(text) else ''
        excerpt = t_prefix + text[t_start:t_end] + t_suffix
        ms = len(t_prefix) + (t_idx - t_start)
        me = ms + search_len
        return excerpt, ms, me
    else:
        return truncate_text(text, 400), -1, -1


def get_match_excerpts(file_path, search_str, case_sensitive=False, context_chars=200, word_app=None):
    """Extract excerpts centered around the search keyword.

    - Keyword always at the center of the excerpt (context_chars before/after)
    - Searches ALL paragraphs (both Chinese and English)
    - Bilingual docs: ZH paragraphs listed first, then EN paragraphs in same order.
      So zh_paragraphs[i] corresponds to en_paragraphs[i].
    - Returns translation_match_start/end for translation highlighting
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".docx":
        paragraphs = get_paragraphs_docx(file_path)
    elif ext == ".doc":
        paragraphs = get_paragraphs_doc(file_path, word_app)
    else:
        return []

    if not paragraphs or not search_str:
        return []

    # Split into zh/en lists while preserving order correspondence
    zh_paragraphs = []
    en_paragraphs = []
    for para in paragraphs:
        if is_english(para):
            en_paragraphs.append(para)
        else:
            zh_paragraphs.append(para)

    excerpts = []
    search_len = len(search_str)

    if case_sensitive:
        target_str = search_str
        match_func = lambda text, target: target in text
        find_func = lambda text, target: text.find(target)
    else:
        target_str = search_str.lower()
        match_func = lambda text, target: target in text.lower()
        find_func = lambda text, target: text.lower().find(target)

    # Helper: build excerpt from a matched paragraph
    def make_excerpt(para, idx):
        match_end_idx = idx + search_len
        excerpt_start = max(0, idx - context_chars)
        excerpt_end = min(len(para), match_end_idx + context_chars)
        prefix = '...' if excerpt_start > 0 else ''
        suffix = '...' if excerpt_end < len(para) else ''
        excerpt_text = prefix + para[excerpt_start:excerpt_end] + suffix
        ms = len(prefix) + (idx - excerpt_start)
        me = ms + search_len
        return excerpt_text, ms, me

    # Search Chinese paragraphs
    for i, zh_para in enumerate(zh_paragraphs):
        if not match_func(zh_para, target_str):
            continue

        idx = find_func(zh_para, target_str)
        excerpt_text, ms, me = make_excerpt(zh_para, idx)

        # Translation: corresponding EN paragraph by index
        translation_text = ""
        translation_match_start = -1
        translation_match_end = -1

        if i < len(en_paragraphs):
            en_para = en_paragraphs[i]
            translation_text, translation_match_start, translation_match_end = \
                _build_translation(en_para, target_str, search_len, match_func, find_func, context_chars)

        excerpts.append({
            'text': excerpt_text,
            'match_start': ms,
            'match_end': me,
            'is_english': False,
            'translation': translation_text,
            'translation_match_start': translation_match_start,
            'translation_match_end': translation_match_end,
        })

    # Search English paragraphs
    for i, en_para in enumerate(en_paragraphs):
        if not match_func(en_para, target_str):
            continue

        idx = find_func(en_para, target_str)
        excerpt_text, ms, me = make_excerpt(en_para, idx)

        # Translation: corresponding ZH paragraph by index
        translation_text = ""
        translation_match_start = -1
        translation_match_end = -1

        if i < len(zh_paragraphs):
            zh_para = zh_paragraphs[i]
            translation_text, translation_match_start, translation_match_end = \
                _build_translation(zh_para, target_str, search_len, match_func, find_func, context_chars)

        excerpts.append({
            'text': excerpt_text,
            'match_start': ms,
            'match_end': me,
            'is_english': True,
            'translation': translation_text,
            'translation_match_start': translation_match_start,
            'translation_match_end': translation_match_end,
        })

    return excerpts


def search_in_file(file_path, search_str, case_sensitive=False, word_app=None):
    text = extract_text(file_path, word_app)
    
    if case_sensitive:
        return search_str in text
    else:
        return search_str.lower() in text.lower()


def search_docs(folder_path, search_str, case_sensitive=False, progress_callback=None):
    results = []
    doc_files = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in (".doc", ".docx"):
                doc_files.append(os.path.join(root, file))

    total = len(doc_files)
    print(f"Found {total} document files to search...")

    if total == 0:
        return results

    # Separate by type: .docx (fast, python-docx) vs .doc (Word COM)
    docx_files = [f for f in doc_files
                  if os.path.splitext(f)[1].lower() == '.docx']
    doc_old_files = [f for f in doc_files
                     if os.path.splitext(f)[1].lower() == '.doc']

    processed = 0

    def _report_progress():
        nonlocal processed
        processed += 1
        if progress_callback:
            progress_callback(processed, total)

    # --- .docx: parallel search with python-docx (fast, no Word needed) ---
    if docx_files:
        from concurrent.futures import ThreadPoolExecutor, as_completed

        def _search_docx(fp):
            try:
                if search_in_file(fp, search_str, case_sensitive, None):
                    return fp
            except Exception as e:
                print(f"Warning: Error processing {fp}: {str(e)}", file=sys.stderr)
            return None

        max_workers = min(4, os.cpu_count() or 1)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_map = {executor.submit(_search_docx, f): f for f in docx_files}
            for future in as_completed(future_map):
                _report_progress()
                result = future.result()
                if result is not None:
                    results.append(result)

    # --- .doc: sequential search with Word COM (slower, needs Word) ---
    word_app = None
    try:
        if doc_old_files:
            if sys.platform != 'win32':
                print("Warning: .doc files require Windows with Microsoft Word "
                      "installed.", file=sys.stderr)
                print("Skipping .doc files...", file=sys.stderr)
            else:
                import win32com.client
                # DispatchEx creates a SEPARATE Word instance to avoid
                # interfering with any Word documents the user has open.
                # (Dispatch would attach to the existing user instance.)
                word_app = win32com.client.DispatchEx("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0  # wdAlertsNone
                try:
                    word_app.ScreenUpdating = False  # big speedup: no UI redraw
                except Exception:
                    pass
                try:
                    word_app.AutomationSecurity = 3  # disable macros
                except Exception:
                    pass
                try:
                    word_app.NormalTemplate.Saved = True  # skip "save Normal?" prompt
                except Exception:
                    pass

            for file_path in doc_old_files:
                _report_progress()
                print(f"Processing doc: {os.path.basename(file_path)}")
                try:
                    if search_in_file(file_path, search_str, case_sensitive, word_app):
                        results.append(file_path)
                except Exception as e:
                    print(f"Warning: Error processing {file_path}: {str(e)}",
                          file=sys.stderr)
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass

    return results


def main():
    parser = argparse.ArgumentParser(description='Search for a string in .doc and .docx files')
    parser.add_argument('search_str', help='The string to search for')
    parser.add_argument('folder_path', help='The folder path to search in')
    parser.add_argument('-c', '--case-sensitive', action='store_true', help='Case sensitive search')
    
    args = parser.parse_args()
    
    if not os.path.isdir(args.folder_path):
        print(f"Error: {args.folder_path} is not a valid directory")
        sys.exit(1)
    
    results = search_docs(args.folder_path, args.search_str, args.case_sensitive)
    
    print("\n" + "="*50)
    if results:
        print(f"Found {len(results)} files containing '{args.search_str}':")
        for result in results:
            print(result)
    else:
        print(f"No files found containing '{args.search_str}'")


if __name__ == "__main__":
    main()