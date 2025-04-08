import os
import uuid
import json
import pandas as pd
import ezdxf
import pinecone
from io import StringIO, BytesIO
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer
import numpy as np

# --- Setup ---
PINECONE_API_KEY = "pcsk_V4zE3_3QKPTNNcTaHE2BXHAPLFjViZsZvRRY3Vhf6hRpDRNb3cgRGpSg1conQdt2VDzSB"
PINECONE_ENV = "us-east-1"
INDEX_NAME = "quotefusion"

pinecone.init(api_key=PINECONE_API_KEY, environment=PINECONE_ENV)
index = pinecone.Index(INDEX_NAME)

# --- Utilities ---
def embed_text(text):
    if model is None:
        raise RuntimeError("Embedding model is not available.")

    return model.encode(text).tolist()

def create_metadata(file_name, content_type, chunk_id):
    return {
        "source": file_name,
        "type": content_type,
        "chunk_id": chunk_id
    }

# --- File Handlers (from file-like objects) ---
def process_csv_file(file_obj, filename):
    df = pd.read_csv(file_obj)
    records = []
    for i, row in df.iterrows():
        text = row.astype(str).str.cat(sep=" ")
        vector = embed_text(text)
        metadata = create_metadata(filename, "csv", i)
        records.append((str(uuid.uuid4()), vector, metadata))
    return records

def process_json_file(file_obj, filename):
    data = json.load(file_obj)
    if isinstance(data, dict):
        data = [data]

    records = []
    for i, item in enumerate(data):
        text = json.dumps(item)
        vector = embed_text(text)
        metadata = create_metadata(filename, "json", i)
        records.append((str(uuid.uuid4()), vector, metadata))
    return records

def process_dxf_file(file_obj, filename):
    buffer = BytesIO(file_obj.read())
    buffer.seek(0)
    doc = ezdxf.read(buffer)
    msp = doc.modelspace()

    records = []
    for i, entity in enumerate(msp):
        if entity.dxftype() == 'LINE':
            text = f"LINE from {entity.dxf.start} to {entity.dxf.end}"
            vector = embed_text(text)
            metadata = create_metadata(filename, "dxf_line", i)
            records.append((str(uuid.uuid4()), vector, metadata))
    return records

# --- GPT File Upload Handler ---
def upload_from_gpt_file(file_obj, filename):
    ext = os.path.splitext(filename)[1].lower()
    if ext == '.csv':
        records = process_csv_file(file_obj, filename)
    elif ext == '.json':
        records = process_json_file(file_obj, filename)
    elif ext == '.dxf':
        records = process_dxf_file(file_obj, filename)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    for i in range(0, len(records), 100):
        index.upsert(vectors=records[i:i+100])
    print(f"Uploaded {len(records)} vectors from {filename}")