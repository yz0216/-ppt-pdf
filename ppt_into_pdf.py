import os
import glob
import comtypes.client

# 讓使用者輸入來源與輸出資料夾
folder_path = input("請輸入 PPT 資料夾路徑：").strip()
output_folder = input("請輸入 PDF 輸出資料夾路徑：").strip()

# 如果輸出資料夾不存在就建立
os.makedirs(output_folder, exist_ok=True)

# 找出所有 ppt 或 pptx 檔案（同時支援大小寫）
ppt_files = glob.glob(os.path.join(folder_path, "*.ppt")) \
           + glob.glob(os.path.join(folder_path, "*.PPT")) \
           + glob.glob(os.path.join(folder_path, "*.pptx")) \
           + glob.glob(os.path.join(folder_path, "*.PPTX"))

print("找到的檔案：", ppt_files)

powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
powerpoint.Visible = 1

for ppt_file in ppt_files:
    # PDF 檔案輸出路徑 → 使用者指定的 output_folder
    base_name = os.path.basename(ppt_file)
    pdf_file = os.path.join(output_folder, base_name.replace(".pptx", ".pdf").replace(".PPTX", ".pdf").replace(".ppt", ".pdf").replace(".PPT", ".pdf"))
    presentation = powerpoint.Presentations.Open(ppt_file)
    presentation.SaveAs(pdf_file, 32)  # 32 = PDF 格式
    presentation.Close()
    print(f"已轉換：{ppt_file} → {pdf_file}")

powerpoint.Quit()
print("全部轉換完成！PDF 檔案都在：", output_folder)
