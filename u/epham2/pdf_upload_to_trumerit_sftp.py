from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)
import wmill
import zipfile
import os
import tempfile
import shutil
import posixpath
import paramiko  # SFTP (SSH) client
import csv
from datetime import date, datetime
import io

# --- PDF/Image libs (required) ---
try:
    from PyPDF2 import PdfMerger  # type: ignore
    from PIL import Image  # type: ignore

    HAS_PDF_LIBS = True
except ImportError:
    PdfMerger = None
    Image = None
    HAS_PDF_LIBS = False


URL = "https://nnas.smapply.io/admin/awards/286165/manage/"
USERNAME = "epham@nnas.ca"
PASSWORD_VAR = "u/epham/SMAPPLY_EDDIE_PASSWORD"

SFTP_HOST = "https://rb-ftp-connect.nnas.ca/web/client/login"
SFTP_PORT = 2022  # adjust if your SFTP uses a non-standard port
SFTP_USER_VAR = "u/knguyen/trumerit_FTP_username"
SFTP_PW_VAR = "u/knguyen/trumerit_ftp_pw"
SFTP_BASE_DIR = "/TO_TRUMERIT"


def log(msg: str, steps: list):
    """Helper to log to stdout and accumulate in steps list."""
    print(msg)
    steps.append(msg)


def download_and_extract_folder_for_reference(page, reference_id: str, steps: list):
    """
    Search by reference_id, open application, select desired download format,
    download ZIP, extract it, and return (folder_path, folder_name, temp_root_dir).
    """
    log(f"[STEP] Searching for application {reference_id}", steps)

    # Search for the application
    search_input = page.locator(
        "xpath=//*[@id='application-grid']/div/div[3]/div[2]/div/div[1]/input"
    )
    search_input.wait_for(state="visible", timeout=30000)
    search_input.fill(reference_id)

    search_button = page.locator(
        "xpath=//*[@id='application-grid']/div/div[3]/div[2]/div/div[1]/span/button"
    )
    search_button.wait_for(state="visible", timeout=30000)
    search_button.click()

    # Wait for the application row link
    app_link = page.locator(
        "xpath=//*[@id='application-grid']/div/div[4]/div[2]/table/tbody/tr/td[4]/div/a"
    )
    app_link.wait_for(state="visible", timeout=30000)
    log("[OK] Search results loaded, application row visible", steps)

    # Click the application link
    app_link.click()
    page.wait_for_load_state("networkidle")
    log("[STEP] Opened application details page", steps)

    # Click the download buttons
    download_button_1 = page.locator(
        "xpath=//*[@id='maincontent']/div[1]/div[1]/div[2]/div[2]/button[2]"
    )
    download_button_1.wait_for(state="visible", timeout=30000)
    download_button_1.click()
    log("[STEP] Clicked first download button (open download options)", steps)

    # Set download format to "Separate PDF for each selected application (with attachments in original format)"
    format_select = page.locator("#id_download_format")
    format_select.wait_for(state="visible", timeout=30000)
    format_select.select_option("as_zip")
    log(
        "[OK] Set download format to 'as_zip' (separate PDF with attachments in original format)",
        steps,
    )

    download_button_2 = page.locator("xpath=//*[@id='download-btn']")
    download_button_2.wait_for(state="visible", timeout=30000)

    # New tab / popup for download
    with page.context.expect_page() as popup_info:
        download_button_2.click()
    download_page = popup_info.value
    download_page.wait_for_load_state("domcontentloaded")
    log("[STEP] Opened download popup page", steps)

    # Click final download button and capture ZIP
    final_download_button = download_page.locator(
        "xpath=//*[@id='sma-js-download-button']"
    )
    final_download_button.wait_for(state="visible", timeout=30000)

    with download_page.expect_download() as dl_info:
        final_download_button.click()

    download = dl_info.value
    zip_path = download.path()  # waits until file is fully downloaded

    if not zip_path:
        raise RuntimeError("Download path not available")

    log(f"[OK] Download completed: {zip_path}", steps)

    # Create a temp directory and extract ZIP there
    temp_root = tempfile.mkdtemp(prefix="smapply_")
    log(f"[STEP] Extracting ZIP into temp dir: {temp_root}", steps)

    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        zip_ref.extractall(temp_root)
        namelist = zip_ref.namelist()

    # Try to find a top-level directory (e.g., NNAS-7319243499_202511240725)
    entries = [os.path.join(temp_root, entry) for entry in os.listdir(temp_root)]
    dir_candidates = [p for p in entries if os.path.isdir(p)]

    if dir_candidates:
        folder_path = dir_candidates[0]
        folder_name = os.path.basename(folder_path)
        log(f"[OK] Found extracted folder: {folder_name}", steps)
        return folder_path, folder_name, temp_root

    # If no directories, try to infer a folder name from ZIP paths
    inferred_name = None
    for name in namelist:
        parts = name.split("/")
        if len(parts) > 1 and parts[0]:
            inferred_name = parts[0]
            break

    if inferred_name:
        candidate = os.path.join(temp_root, inferred_name)
        if not os.path.isdir(candidate):
            os.makedirs(candidate, exist_ok=True)
            for entry in os.listdir(temp_root):
                if entry == inferred_name:
                    continue
                src = os.path.join(temp_root, entry)
                dst = os.path.join(candidate, entry)
                shutil.move(src, dst)
        folder_path = candidate
        folder_name = inferred_name
        log(f"[OK] Inferred and prepared folder: {folder_name}", steps)
    else:
        # Fallback: treat the extraction root as the folder, use reference_id as folder name
        folder_path = temp_root
        folder_name = reference_id
        log(
            f"[WARN] No folder found; using temp root as folder with name={folder_name}",
            steps,
        )

    return folder_path, folder_name, temp_root


def count_files_in_folder(folder_path: str) -> int:
    """
    Count total number of files (all types) in folder_path recursively.
    (Folders themselves are not counted.)
    """
    count = 0
    for _root, _dirs, files in os.walk(folder_path):
        count += len(files)
    return count


def remove_pdf_odd_files(folder_path: str, steps: list):
    """
    Remove any file in folder_path (recursively) whose name contains '.pdf_'.
    Returns a list of removed file paths (relative to folder_path).
    """
    log(f"[STEP] Removing files containing '.pdf_' under {folder_path}", steps)
    removed = []
    for root, dirs, files in os.walk(folder_path):
        for fname in files:
            if ".pdf_" in fname:
                full_path = os.path.join(root, fname)
                rel_path = os.path.relpath(full_path, folder_path)
                try:
                    os.remove(full_path)
                    removed.append(rel_path)
                    log(f"[OK] Removed file: {rel_path}", steps)
                except Exception as e:
                    removed.append(f"{rel_path} (failed to remove: {e})")
                    log(f"[WARN] Failed to remove {rel_path}: {e}", steps)
    if not removed:
        log("[OK] No '.pdf_' files found to remove", steps)
    return removed


def find_main_pdf(folder_path: str, reference_id: str, steps: list):
    """
    Find the 'main' PDF whose filename contains the reference_id.
    Returns its full path or None if not found.
    """
    candidates = []
    for root, dirs, files in os.walk(folder_path):
        for fname in files:
            if fname.lower().endswith(".pdf") and reference_id in fname:
                full_path = os.path.join(root, fname)
                candidates.append(full_path)

    if not candidates:
        log(f"[WARN] No main PDF found containing reference_id '{reference_id}'", steps)
        return None

    # If multiple, just pick the first one
    main_pdf = candidates[0]
    rel_main = os.path.relpath(main_pdf, folder_path)
    log(
        f"[OK] Identified main PDF (excluded from timestamp rename & attachments merge): {rel_main}",
        steps,
    )
    return main_pdf


def rename_files_with_timestamp(
    folder_path: str,
    timestamp: str,
    main_pdf_path: str | None,
    steps: list,
):
    """
    Append the same timestamp to all files:
      - Non-PDF files: <name>_<timestamp><ext>
      - PDF files:     <name>_<timestamp>.pdf
    EXCEPT the main_pdf_path (if provided), which is left unchanged.
    """
    log(f"[STEP] Renaming files with timestamp {timestamp}", steps)

    for root, dirs, files in os.walk(folder_path):
        for fname in files:
            old_path = os.path.join(root, fname)

            # Skip the main reference PDF (if identified)
            if main_pdf_path and os.path.abspath(old_path) == os.path.abspath(
                main_pdf_path
            ):
                rel = os.path.relpath(old_path, folder_path)
                log(f"[OK] Skipping timestamp rename for main PDF: {rel}", steps)
                continue

            base, ext = os.path.splitext(fname)
            new_name = f"{base}_{timestamp}{ext}"
            new_path = os.path.join(root, new_name)

            try:
                os.rename(old_path, new_path)
                rel_old = os.path.relpath(old_path, folder_path)
                rel_new = os.path.relpath(new_path, folder_path)
                log(f"[OK] Renamed file: {rel_old} -> {rel_new}", steps)
            except Exception as e:
                log(f"[WARN] Failed to rename {old_path}: {e}", steps)


def create_combined_attachments_pdf(
    folder_path: str,
    reference_id: str,
    timestamp: str,
    main_pdf_path: str | None,
    steps: list,
):
    """
    Create a combined PDF of all attachments:
      - All PDFs EXCEPT main_pdf_path (if given)
      - All images (png/jpg/jpeg/etc) as converted PDF pages
    Writes combined PDF into folder_path as:
      '<reference_id>_attachments_combined_<timestamp>.pdf'

    Returns:
      (combined_pdf_path or None, list_of_attachment_source_paths)
    """
    log("[STEP] Preparing to create combined attachments PDF", steps)

    if not HAS_PDF_LIBS:
        raise RuntimeError(
            "PyPDF2/Pillow are required to create combined attachments PDF but are not installed."
        )

    # Collect attachment candidates (PDFs and images) excluding main PDF
    attachment_files = []
    image_exts = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}

    for root, dirs, files in os.walk(folder_path):
        for fname in files:
            full_path = os.path.join(root, fname)
            # Skip the main PDF
            if main_pdf_path and os.path.abspath(full_path) == os.path.abspath(
                main_pdf_path
            ):
                continue

            ext = os.path.splitext(fname)[1].lower()
            if ext == ".pdf" or ext in image_exts:
                attachment_files.append(full_path)

    if not attachment_files:
        log("[WARN] No attachment files (PDF/image) found to combine", steps)
        return None, []

    log(f"[STEP] Found {len(attachment_files)} attachment file(s) to combine", steps)

    merger = PdfMerger()
    tmp_img_pdf_dir = tempfile.mkdtemp(prefix="imgpdf_")
    tmp_img_pdfs = []

    combined_path = None

    try:
        for path in attachment_files:
            ext = os.path.splitext(path)[1].lower()
            try:
                if ext == ".pdf":
                    rel = os.path.relpath(path, folder_path)
                    log(f"[STEP] Adding PDF to merger: {rel}", steps)
                    merger.append(path)
                else:
                    # Convert image to temp PDF
                    rel = os.path.relpath(path, folder_path)
                    log(f"[STEP] Converting image to PDF and adding: {rel}", steps)
                    img = Image.open(path)
                    if img.mode != "RGB":
                        img = img.convert("RGB")
                    tmp_pdf_path = os.path.join(
                        tmp_img_pdf_dir,
                        os.path.basename(path) + ".pdf",
                    )
                    img.save(tmp_pdf_path, "PDF")
                    tmp_img_pdfs.append(tmp_pdf_path)
                    merger.append(tmp_pdf_path)
            except Exception as e:
                rel = os.path.relpath(path, folder_path)
                log(f"[WARN] Skipped {rel} due to error while merging: {e}", steps)

        if len(merger.pages) == 0:
            log(
                "[WARN] No pages were added to the attachments merger; skipping combined PDF creation",
                steps,
            )
            return None, attachment_files

        combined_name = f"{reference_id}_attachments_combined_{timestamp}.pdf"
        combined_path = os.path.join(folder_path, combined_name)
        merger.write(combined_path)
        merger.close()
        rel_combined = os.path.relpath(combined_path, folder_path)
        log(f"[OK] Created combined attachments PDF: {rel_combined}", steps)
        return combined_path, attachment_files

    finally:
        # Clean up temporary image PDFs
        shutil.rmtree(tmp_img_pdf_dir, ignore_errors=True)


def ensure_remote_dirs_sftp(sftp, remote_dir: str, steps: list):
    """
    Ensure that remote_dir exists on the SFTP server, creating directories as needed.
    Uses absolute POSIX paths.
    """
    remote_dir = "/" + remote_dir.strip("/")
    parts = remote_dir.strip("/").split("/")
    path = ""
    for part in parts:
        path = path + "/" + part
        try:
            sftp.chdir(path)
        except IOError:
            log(f"[STEP] Creating remote dir: {path}", steps)
            sftp.mkdir(path)
            sftp.chdir(path)


def upload_folder_and_csv_to_sftp(
    folder_path: str,
    folder_name: str,
    csv_filename: str,
    csv_row: list,
    sftp_user: str,
    sftp_pw: str,
    steps: list,
):
    """
    Using a single SFTP connection:
      - Upload each file in folder_path to /TO_TRUMERIT/<folder_name>/ (preserving structure)
      - Append the row to /TO_TRUMERIT/<csv_filename> if it exists,
        or create it if not.
    Returns (remote_root, uploaded_files, csv_remote_path)
    """
    uploaded = []

    log(f"[STEP] Connecting to SFTP {SFTP_HOST}:{SFTP_PORT} as {sftp_user}", steps)
    transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
    try:
        transport.connect(username=sftp_user, password=sftp_pw)
        sftp = paramiko.SFTPClient.from_transport(transport)
        log("[OK] SFTP login successful (SSH)", steps)

        # 1. Upload folder files to /TO_TRUMERIT/<folder_name>/
        remote_root = posixpath.join(SFTP_BASE_DIR, folder_name)
        log(f"[STEP] Ensuring remote folder root: {remote_root}", steps)
        ensure_remote_dirs_sftp(sftp, remote_root, steps)

        log(
            f"[STEP] Uploading files from {folder_path} to SFTP under {remote_root}",
            steps,
        )
        for root, dirs, files in os.walk(folder_path):
            for fname in files:
                local_path = os.path.join(root, fname)
                rel_path = os.path.relpath(local_path, folder_path)
                rel_path_posix = rel_path.replace(os.sep, "/")
                remote_path = posixpath.join(remote_root, rel_path_posix)

                remote_dir = posixpath.dirname(remote_path)
                ensure_remote_dirs_sftp(sftp, remote_dir, steps)

                log(f"[STEP] Uploading {rel_path_posix} -> {remote_path}", steps)
                sftp.put(local_path, remote_path)
                uploaded.append(remote_path)
                log(f"[OK] Uploaded {remote_path}", steps)

        # 2. Append row to /TO_TRUMERIT/<YYYY-MM-DD>.csv (or create it)
        csv_remote_path = posixpath.join(SFTP_BASE_DIR, csv_filename)

        # Serialize just the new row to CSV line
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(csv_row)
        line = buf.getvalue()

        # Check if remote CSV exists
        try:
            sftp.stat(csv_remote_path)
            exists = True
            log(f"[STEP] Remote CSV {csv_remote_path} exists; appending row", steps)
        except IOError:
            exists = False
            log(
                f"[STEP] Remote CSV {csv_remote_path} does not exist; creating new file",
                steps,
            )

        mode = "ab" if exists else "wb"
        with sftp.open(csv_remote_path, mode) as f:
            f.write(line.encode("utf-8"))
        log(f"[OK] Wrote row to remote CSV {csv_remote_path}: {csv_row}", steps)

        log("[OK] SFTP uploads (files + CSV row) complete", steps)
        return remote_root, uploaded, csv_remote_path

    finally:
        try:
            transport.close()
            log("[OK] SFTP connection closed", steps)
        except Exception:
            pass


def extract_rn_from_additional(additional_content: str, steps: list) -> str:
    """
    Extract RN as: string after the 7th comma before the next comma.
    Example: "..., , , , , , , LPN, email" -> "LPN"
    """
    if not additional_content:
        log("[WARN] additional_content is empty; RN left empty", steps)
        return ""

    parts = additional_content.split(",")
    if len(parts) >= 8:
        rn_candidate = parts[7].strip()
        log(f"[OK] Extracted RN='{rn_candidate}' from additional_content", steps)
        return rn_candidate
    else:
        log(
            "[WARN] additional_content has fewer than 8 comma-separated fields; RN left empty",
            steps,
        )
        return ""


def extract_email_from_additional(additional_content: str, steps: list) -> str:
    """
    Extract email as: string after the 8th comma before the next comma.
    Example: "..., LPN,nnas.test.email3@gmail.com " -> "nnas.test.email3@gmail.com"
    """
    if not additional_content:
        log("[WARN] additional_content is empty; email left empty", steps)
        return ""

    parts = additional_content.split(",")
    if len(parts) >= 9:
        email_candidate = parts[8].strip()
        log(f"[OK] Extracted email='{email_candidate}' from additional_content", steps)
        return email_candidate
    else:
        log(
            "[WARN] additional_content has fewer than 9 comma-separated fields; email left empty",
            steps,
        )
        return ""


def build_csv_row(
    user: str,
    email: str,
    user_id: int,
    reference_id: str,
    rn: str,
    file_count: int,
    steps: list,
):
    """
    Build the final CSV row:
    user, email, user_id, reference_id, RN, file_count
    """
    user_id_str = "" if user_id is None else str(user_id)
    row = [
        user or "",
        email or "",
        user_id_str,
        reference_id or "",
        rn or "",
        str(file_count),
    ]
    log(f"[STEP] Built CSV row: {row}", steps)
    return row


def main(
    reference_id: str = "",
    user: str = "",
    user_id: int = None,
    additional_content: str = "",
    trigger: dict = None,
):
    """
    Windmill entrypoint.

    Final CSV row format:
      user, email_from_additional, user_id, reference_id, RN, file_count

    Flow:
      - REFUSE to run if PyPDF2/Pillow are not installed
      - Merge values from 'trigger' into parameters (if provided)
      - Login to SMApply, download ZIP (as 'as_zip'), extract
      - Clean up filenames:
          * Remove any '.pdf_' files
          * Generate a single timestamp in main
          * Find main PDF (filename contains reference_id)
          * Rename ALL files with that timestamp, except the main PDF
          * Create combined attachments PDF with the same timestamp
          * Move all attachments used for the combined PDF + the combined PDF
            into a subfolder: <reference_id>_attachments_<timestamp>/
      - Count files in folder (files in subfolders included; folders not counted)
      - Parse RN and email from additional_content
      - Build CSV row (with file_count)
      - Append row to local YYYY-MM-DD.csv
      - Single SFTP connection:
          * Upload each file to /TO_TRUMERIT/<folder_name>/ (including attachments subfolder)
          * Append row to /TO_TRUMERIT/<YYYY-MM-DD>.csv (or create if missing)
      - Clean up temp folder
    """
    steps = []

    # Hard requirement: PyPDF2 & Pillow must be installed
    if not HAS_PDF_LIBS:
        msg = "PyPDF2 and Pillow are required but not installed; aborting script."
        log(f"[ERROR] {msg}", steps)
        return {"status": "error", "message": msg, "steps": steps}

    # Merge from trigger JSON if provided
    if trigger is not None and isinstance(trigger, dict):
        log(f"[DEBUG] Raw trigger payload: {trigger}", steps)
        if not reference_id:
            reference_id = trigger.get("reference_id", reference_id)
        if not user:
            user = trigger.get("user", user)
        if user_id is None and "user_id" in trigger:
            user_id = trigger.get("user_id")
        if not additional_content:
            additional_content = trigger.get("additional_content", additional_content)

    log(
        f"[DEBUG] Resolved inputs -> user={user!r}, user_id={user_id!r}, "
        f"reference_id={reference_id!r}, additional_content={additional_content!r}",
        steps,
    )

    if not reference_id:
        msg = "reference_id is required but missing"
        log(f"[ERROR] {msg}", steps)
        return {"status": "error", "message": msg, "steps": steps}

    password = wmill.get_variable(PASSWORD_VAR)
    if not password:
        msg = f"Password variable {PASSWORD_VAR} is empty or missing"
        log(f"[ERROR] {msg}", steps)
        return {"status": "error", "message": msg, "steps": steps}

    sftp_user = wmill.get_variable(SFTP_USER_VAR)
    sftp_pw = wmill.get_variable(SFTP_PW_VAR)
    if not sftp_user or not sftp_pw:
        msg = "SFTP username/password variables are empty or missing"
        log(f"[ERROR] {msg}", steps)
        return {"status": "error", "message": msg, "steps": steps}

    log(f"[START] Processing reference_id={reference_id}", steps)

    temp_root = None
    folder_path = None
    folder_name = None

    try:
        # 1) Playwright: login + download + extract
        with sync_playwright() as p:
            log("[STEP] Launching Chromium", steps)
            browser = p.chromium.launch(
                executable_path="/usr/bin/chromium",  # adjust if needed
                args=[
                    "--no-sandbox",
                    "--single-process",
                    "--no-zygote",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                ],
                headless=True,
            )
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            log(f"[STEP] Navigating to login URL: {URL}", steps)
            page.goto(URL, wait_until="domcontentloaded")

            log("[STEP] Filling email", steps)
            page.wait_for_selector("#email-field", timeout=15000)
            page.fill("#email-field", USERNAME)

            page.keyboard.press("Enter")
            try:
                page.wait_for_selector("#password-field", timeout=8000)
                log("[OK] Password field appeared after Enter", steps)
            except PlaywrightTimeoutError:
                log(
                    "[WARN] Password field did not appear after Enter, clicking 'Next' button",
                    steps,
                )
                next_btn = page.get_by_role("button", name="Next")
                next_btn.click()
                page.wait_for_selector("#password-field", timeout=15000)
                log("[OK] Password field appeared after clicking 'Next'", steps)

            log("[STEP] Filling password", steps)
            page.fill("#password-field", password)

            page.keyboard.press("Enter")
            try:
                page.wait_for_url("**/admin/awards/286165/manage/**", timeout=20000)
                log("[OK] Reached manage page after Enter", steps)
            except PlaywrightTimeoutError:
                log(
                    "[WARN] Manage page not reached after Enter, clicking 'Log in' button",
                    steps,
                )
                login_btn = page.get_by_role("button", name="Log in")
                login_btn.click()
                page.wait_for_url("**/admin/awards/286165/manage/**", timeout=20000)
                log("[OK] Reached manage page after clicking 'Log in'", steps)

            folder_path, folder_name, temp_root = (
                download_and_extract_folder_for_reference(page, reference_id, steps)
            )

            # 2) Remove any '.pdf_' files first
            removed_files = remove_pdf_odd_files(folder_path, steps)

            # 3) Generate a single timestamp in main and reuse everywhere
            timestamp = datetime.utcnow().strftime("%Y%m%d%H%M%S")
            log(f"[STEP] Using unified timestamp for all renames: {timestamp}", steps)

            # 4) Identify main PDF (reference_id in filename), if present
            main_pdf_path = find_main_pdf(folder_path, reference_id, steps)

            # 5) Rename all files with that timestamp, except the main PDF
            rename_files_with_timestamp(folder_path, timestamp, main_pdf_path, steps)

                       # 6) Create combined attachments PDF (with same timestamp in filename)
            combined_path, attachment_files = create_combined_attachments_pdf(
                folder_path, reference_id, timestamp, main_pdf_path, steps
            )

            # 7) If we created a combined PDF, move all attachments (but NOT the combined PDF)
            #    into a dedicated subfolder: <reference_id>_attachments_<timestamp>
            if combined_path:
                attachments_folder_name = f"{reference_id}_attachments_{timestamp}"
                attachments_folder_path = os.path.join(folder_path, attachments_folder_name)
                os.makedirs(attachments_folder_path, exist_ok=True)
                log(f"[STEP] Moving attachments into subfolder: {attachments_folder_name}", steps)

                # Move original attachment sources into the attachments folder
                for src in attachment_files:
                    if not os.path.exists(src):
                        continue
                    dest = os.path.join(attachments_folder_path, os.path.basename(src))
                    try:
                        shutil.move(src, dest)
                        rel_src = os.path.relpath(src, folder_path)
                        rel_dest = os.path.relpath(dest, folder_path)
                        log(f"[OK] Moved attachment: {rel_src} -> {rel_dest}", steps)
                    except Exception as e:
                        log(f"[WARN] Failed to move attachment {src}: {e}", steps)

                # IMPORTANT: leave combined PDF in the main folder, next to the main/reference PDF
                rel_combined = os.path.relpath(combined_path, folder_path)
                log(f"[OK] Combined PDF remains in main folder: {rel_combined}", steps)

            browser.close()
            log("[OK] Browser closed", steps)


        # 8) Count files in folder (including subfolders; folders themselves are not counted)
        file_count = count_files_in_folder(folder_path)
        log(f"[STEP] Counted {file_count} files in folder {folder_path}", steps)

        # 9) Parse RN and email from additional_content
        rn = extract_rn_from_additional(additional_content, steps)
        email = extract_email_from_additional(additional_content, steps)

        # 10) Build CSV row
        csv_row = build_csv_row(
            user, email, user_id, reference_id, rn, file_count, steps
        )

        # 11) Append to local daily CSV
        today_str = date.today().isoformat()
        root_dir = os.getcwd()
        csv_filename = f"{today_str}.csv"
        csv_path = os.path.join(root_dir, csv_filename)
        log(f"[STEP] Updating local daily CSV: {csv_path}", steps)

        file_exists = os.path.exists(csv_path)
        mode = "a" if file_exists else "w"
        with open(csv_path, mode, newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(csv_row)

        if file_exists:
            log(f"[OK] Appended row to existing local CSV: {csv_row}", steps)
        else:
            log(f"[OK] Created new local CSV and wrote first row: {csv_row}", steps)

        # 12) Single SFTP connection for BOTH files + CSV row
        remote_root, uploaded_files, csv_remote_path = upload_folder_and_csv_to_sftp(
            folder_path, folder_name, csv_filename, csv_row, sftp_user, sftp_pw, steps
        )

        # 13) Cleanup temp directory
        if temp_root and os.path.isdir(temp_root):
            shutil.rmtree(temp_root, ignore_errors=True)
            log(f"[OK] Cleaned up temp directory {temp_root}", steps)

        log("[DONE] Workflow completed successfully", steps)
        return {
            "status": "ok",
            "reference_id": reference_id,
            "local_folder_name": folder_name,
            "file_count": file_count,
            "sftp_remote_root": remote_root,
            "removed_files": removed_files,
            "uploaded_files": uploaded_files,
            "csv_file": csv_path,
            "csv_row": csv_row,
            "csv_remote_path": csv_remote_path,
            "steps": steps,
        }

    except PlaywrightTimeoutError as e:
        if temp_root and os.path.isdir(temp_root):
            shutil.rmtree(temp_root, ignore_errors=True)
            log(f"[OK] Cleaned up temp directory {temp_root}", steps)
        msg = f"timeout while processing reference {reference_id}: {e}"
        log(f"[ERROR] {msg}", steps)
        return {"status": "error", "message": msg, "steps": steps}

    except Exception as e:
        if temp_root and os.path.isdir(temp_root):
            shutil.rmtree(temp_root, ignore_errors=True)
            log(f"[OK] Cleaned up temp directory {temp_root}", steps)
        msg = f"unexpected error: {e}"
        log(f"[ERROR] {msg}", steps)
        return {"status": "error", "message": msg, "steps": steps}
