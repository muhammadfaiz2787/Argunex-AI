FROM python:3.10-slim

WORKDIR /code

# Salin file requirements ke dalam container
COPY ./requirements.txt /code/requirements.txt

# Instal semua dependensi pustaka Python
RUN pip install --no-cache-dir --upgrade -r /code/requirements.txt

# Salin semua kode backend (termasuk app.py dan aset lainnya)
COPY . .

# Jalankan server FastAPI dengan Uvicorn di port 7860 (Port wajib Hugging Face)
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "7860"]