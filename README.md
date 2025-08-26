# Master Data Log Book Generator — Yang Lab

This tool creates a master Excel file from the **Human Data** folder in OneDrive.  
The Excel contains:
- Number of DICOM files (`.dcm` and `.ima`)  
- Number of Twix files (`.dat`)  
- Number of subfolders inside each subject folder  
- Acquisition date (from DICOM tag)  
- File creation date  
- Reconstruction (%)  

---

## How to Run
1. Download the file **CreateMasterExcel.exe** from this repository.  
2. Double-click the file.  
3. Wait until the black window shows:  
   - *Processing starts…*  
   - *Master Excel saved at: …*  
4. Open the Excel file from the saved location.  

---

## Notes
- Works for **Human Data folder in Onedrive only** (Animal Data version coming later).  
- If Windows gives a security warning, choose **More info → Run anyway**.  
# Master-Data-Log-Book-Generator-YANG-LAB
Automated tool for Yang Lab to scan OneDrive subject folders, count DICOM/Twix files, extract acquisition dates, and generate a master Excel log book.
