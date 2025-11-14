import pandas as pd
import re
import zipfile
import os
import shutil

# 1) 엑셀 파일 로드 (.xlsx 정상 지원)
excel_path = r"C:\Users\ST\Desktop\phone_number\hi.xlsx"
df = pd.read_excel(excel_path, dtype=str)

# 고객번호(3열)를 key로, PDS제목(2열)을 value로 매핑
mapping_title = dict(zip(df.iloc[:, 2], df.iloc[:, 1]))

# 2) ZIP 안의 mp3 파일 검사
zip_path = r"C:\Users\ST\Desktop\phone_number\record_file.zip"
output = []

with zipfile.ZipFile(zip_path, 'r') as z:
    for file in z.namelist():
        if file.lower().endswith(".mp3"):

            match = re.search(r"(010\d{8})", file)
            if not match:
                continue

            phone = match.group(1)
            pds_title = mapping_title.get(phone)

            if pds_title:
                output.append({
                    "PDS제목": pds_title,
                    "고객번호": phone,
                    "mp3파일명": file
                })

# 루트 결과 폴더
result_root = r"C:\Users\ST\Desktop\phone_number\result"

df_out = pd.DataFrame(output)

# 3) PDS제목별 폴더 생성 + mp3 파일 + 엑셀 동시 저장
with zipfile.ZipFile(zip_path, 'r') as z:
    for title, group_df in df_out.groupby("PDS제목"):

        # 폴더 안전 이름 처리
        safe_title = re.sub(r'[\\/*?:"<>|]', "", title)

        folder_path = os.path.join(result_root, safe_title)
        os.makedirs(folder_path, exist_ok=True)

        # 3-1) 엑셀 저장
        excel_path = os.path.join(folder_path, "mp3_매칭결과.xlsx")
        group_df[["PDS제목", "고객번호"]].to_excel(excel_path, index=False)

        # 3-2) mp3 파일 추출해서 폴더 안으로 저장
        for mp3_name in group_df["mp3파일명"]:
    
    # 번호 다시 추출
            match = re.search(r"(010\d{8})", mp3_name)
            if not match:
                continue
            
            phone_num = match.group(1)
            
            # 저장 파일명 → 전화번호 + ".mp3" ONLY
            save_name = f"{phone_num}.mp3"
            mp3_path = os.path.join(folder_path, save_name)

            # 이미 저장되었다면 overwrite 안 함 (중복 방지)
            if os.path.exists(mp3_path):
                continue

            # ZIP에서 원본 파일 읽어서 전화번호.mp3로 저장
            with z.open(mp3_name) as source, open(mp3_path, "wb") as target:
                shutil.copyfileobj(source, target)

print("완료! → PDS제목별 폴더에 엑셀 + mp3 파일까지 모두 저장됨.")
