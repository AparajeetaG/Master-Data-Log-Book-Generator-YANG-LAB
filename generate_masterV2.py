import os
import sys
import time
from datetime import datetime
import pandas as pd
import pydicom

# ----------------- CONFIG -----------------
# Root folder to scan
root_folder = r"C:\Users\guhaa2\OneDrive - Cedars-Sinai Health System\Yang, Hsin-Jung's files - Data_Human"

# Excel save location (OneDrive root)
excel_path = os.path.join(os.path.expanduser("~"), "OneDrive - Cedars-Sinai Health System", "master.xlsx")
# ------------------------------------------


def extract_study_date(dicom_file):
    """Extract StudyDate (YYYYMMDD) from a DICOM file, if present."""
    try:
        ds = pydicom.dcmread(dicom_file, stop_before_pixels=True, force=True)
        if "StudyDate" in ds:
            return ds.StudyDate
    except Exception:
        return None
    return None


def main():
    start_ts = time.time()
    print("=== Master Excel Generator V2 ===")
    print(f"Start time          : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Scanning root       : {root_folder}")
    print(f"Will save Excel to  : {excel_path}")
    print("Processing..."); sys.stdout.flush()

    # ---------- Precompute number of subfolders per subject ----------
    print("Step 1/2: Precomputing number of subfolders per subject..."); sys.stdout.flush()
    subject_subfolder_counts = {}
    for root, dirs, files in os.walk(root_folder):
        parts = root.replace(root_folder, "").strip(os.sep).split(os.sep)
        if len(parts) == 2:  # subject folder level
            subject_folder = parts[1]
            subject_subfolder_counts[subject_folder] = len(dirs)

    # ---------- Walk tree and build rows ----------
    print("Step 2/2: Scanning files and building rows..."); sys.stdout.flush()
    data = []
    files_seen = 0
    rows_built = 0

    for root, dirs, files in os.walk(root_folder):
        if not files:
            continue

        parts = root.replace(root_folder, "").strip(os.sep).split(os.sep)
        if len(parts) < 2:
            continue

        root_folder_name = parts[0]
        subject_folder_name = parts[1]
        subfolder_name = os.sep.join(parts[2:]) if len(parts) > 2 else ""

        # Number of subfolders for this subject
        num_subfolders_in_subject = subject_subfolder_counts.get(subject_folder_name, 0)

        # Counts
        dicom_files = [f for f in files if f.lower().endswith(".dcm") or f.lower().endswith(".ima")]
        twix_files  = [f for f in files if f.lower().endswith(".dat")]

        dicom_count = len(dicom_files)
        twix_count  = len(twix_files)

        # Acquisition date + file creation date from one DICOM (if any)
        study_date = None
        file_creation_date = None
        if dicom_files:
            sample_dicom = os.path.join(root, dicom_files[0])
            study_date_raw = extract_study_date(sample_dicom)
            if study_date_raw:
                try:
                    study_date = datetime.strptime(study_date_raw, "%Y%m%d").strftime("%Y-%m-%d")
                except Exception:
                    study_date = study_date_raw  # keep raw if parse fails

            try:
                file_creation_date = datetime.fromtimestamp(os.path.getctime(sample_dicom)).strftime("%Y-%m-%d")
            except Exception:
                file_creation_date = None

        # Record rows (one per .dat file; if none, still one row)
        if twix_files:
            for dat_file in twix_files:
                data.append([
                    root_folder_name,
                    subject_folder_name,
                    num_subfolders_in_subject,
                    subfolder_name,
                    dicom_count,
                    twix_count,
                    study_date,
                    file_creation_date,
                    dat_file,
                    None  # Reconstruction (%)
                ])
                rows_built += 1
        else:
            data.append([
                root_folder_name,
                subject_folder_name,
                num_subfolders_in_subject,
                subfolder_name,
                dicom_count,
                twix_count,
                study_date,
                file_creation_date,
                None,
                None  # Reconstruction (%)
            ])
            rows_built += 1

        files_seen += len(files)
        if files_seen and files_seen % 5000 == 0:
            print(f"  Progress: scanned ~{files_seen:,} files, built {rows_built:,} rows..."); sys.stdout.flush()

    # ---------- Save to Excel ----------
    print("Saving Excel..."); sys.stdout.flush()
    df = pd.DataFrame(data, columns=[
        "Root Folder",
        "Subject Folder",
        "Number of Subfolders in Subject Folder",
        "Subfolder Name (inside Subject)",
        "Dicom Count",
        "Twix Count",
        "Acquisition Date",
        "File Creation Date",
        "DAT File",
        "Reconstruction (%)"
    ])

    # Ensure target directory exists
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)
    df.to_excel(excel_path, index=False)

    elapsed = time.time() - start_ts
    mm, ss = divmod(int(elapsed), 60)
    print("=== Done ===")
    print(f"End time            : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Total rows written  : {len(df):,}")
    print(f"Excel saved to      : {excel_path}")
    print(f"Elapsed time        : {mm}m {ss}s")
    sys.stdout.flush()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR:", str(e))
        raise
    finally:
        # Keep window open if launched by double-click
        try:
            if not sys.stdin.closed:
                input("Press Enter to exit...")
        except Exception:
            pass
