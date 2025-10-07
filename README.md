# OCR → Financials Excel

แอป Streamlit สำหรับแปลง PDF/รูปภาพงบการเงินไทยเป็น Excel พร้อม OCR และการจัดตารางอัตโนมัติ

## วิธีใช้งาน

1. ติดตั้งไลบรารี Python ที่จำเป็น:
    ```
    pip install -r requirements.txt
    ```
2. (Linux) ติดตั้งแพ็กเกจระบบ:
    ```
    sudo apt-get install -y poppler-utils tesseract-ocr
    ```
3. รันแอป:
    ```
    streamlit run app.py
    ```

## ฟีเจอร์

- อัปโหลด PDF หรือรูปภาพ (PNG/JPG)
- OCR ด้วย OpenTyphoon API หรือ Tesseract (ออฟไลน์)
- แปลงตารางงบการเงินเป็น DataFrame
- ตรวจสอบความสมดุลของงบการเงิน
- ดาวน์โหลดไฟล์ Excel ที่ได้

## หมายเหตุ

- หากใช้ Tesseract ให้ติดตั้งภาษาไทยด้วย `sudo apt-get install tesseract-ocr-tha`
- ต้องมี API Key หากใช้ OpenTyphoon API
