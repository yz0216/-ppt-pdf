import os
import glob
import comtypes.client

# Prompt user for source and output folder paths
folder_path = input("Please enter the PPT folder path: ").strip()
output_folder = input("Please enter the PDF output folder path: ").strip()

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Find all ppt or pptx files (case-insensitive)
ppt_files = glob.glob(os.path.join(folder_path, "*.ppt")) \
            + glob.glob(os.path.join(folder_path, "*.PPT")) \
            + glob.glob(os.path.join(folder_path, "*.pptx")) \
            + glob.glob(os.path.join(folder_path, "*.PPTX"))

print("Files found:", ppt_files)

powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
powerpoint.Visible = 1

for ppt_file in ppt_files:
    # PDF output path -> user-specified output_folder
    base_name = os.path.basename(ppt_file)
    pdf_file = os.path.join(output_folder, base_name.replace(".pptx", ".pdf").replace(".PPTX", ".pdf").replace(".ppt", ".pdf").replace(".PPT", ".pdf"))
    presentation = powerpoint.Presentations.Open(ppt_file)
    presentation.SaveAs(pdf_file, 32)  # 32 = PDF format
    presentation.Close()
    print(f"Converted: {ppt_file} -> {pdf_file}")

powerpoint.Quit()
print("All conversions completed! PDF files are located at:", output_folder)
