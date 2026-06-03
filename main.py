from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from PIL import Image
from ultralytics import YOLO
import io

app = FastAPI(title="Nutrition AI Backend")

# Inisialisasi model YOLOv8 (Otomatis download saat pertama kali dirun)
# Catatan: Nanti file ini diganti dengan model hasil training tabel gizi Anda sendiri
model = YOLO("yolov8n.pt")

@app.get("/")
def home():
    return {"status": "active", "message": "Server AI YOLOv8 Siap!"}

@app.post("/api/v1/scan")
async def scan_nutrition_table(file: UploadFile = File(...)):
    try:
        # 1. Baca file gambar dari HP
        request_object_content = await file.read()
        img = Image.open(io.BytesIO(request_object_content))
        
        # 2. Jalankan deteksi objek AI YOLOv8 langsung pada objek gambar
        results = model(img)
        
        # 3. Ambil data koordinat kotak pembatas (Bounding Boxes) yang ditemukan
        detected_objects = []
        for box in results[0].boxes:
            xyxy = box.xyxy[0].tolist()  # Koordinat [Kiri, Atas, Kanan, Bawah] dalam piksel
            confidence = float(box.conf[0])  # Tingkat keyakinan AI (0.0 sampai 1.0)
            class_id = int(box.cls[0])       # ID Kategori objek
            class_name = model.names[class_id] # Nama kategori (misal: bottle, cup, dll)
            
            detected_objects.append({
                "object_type": class_name,
                "confidence": round(confidence, 2),
                "bounding_box": [round(coord, 1) for coord in xyxy]
            })
        
        return JSONResponse(status_code=200, content={
            "message": "Gambar berhasil diproses oleh YOLOv8!",
            "total_detected": len(detected_objects),
            "objects": detected_objects
        })
        
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Gagal memproses AI: {str(e)}"})