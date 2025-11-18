import os
from azure.storage.blob import BlobServiceClient
from utils import get_env, logger

BLOB_CONN = get_env("AZURE_BLOB_CONN", required=True)
GENERATED_CONTAINER = get_env("GENERATED_CONTAINER", "generated-presentations")

def upload_ppt_to_blob(file_path, file_name):
    blob_service = BlobServiceClient.from_connection_string(BLOB_CONN)
    container_client = blob_service.get_container_client(GENERATED_CONTAINER)
    try:
        container_client.create_container()
    except Exception:
        pass
    with open(file_path, "rb") as data:
        container_client.upload_blob(name=file_name, data=data, overwrite=True)
    logger.info(f"Uploaded generated PPT to Azure Blob: {GENERATED_CONTAINER}/{file_name}")
    return f"{GENERATED_CONTAINER}/{file_name}"

def upload_json_to_blob(json_bytes, blob_name):
    blob_service = BlobServiceClient.from_connection_string(BLOB_CONN)
    container_client = blob_service.get_container_client(GENERATED_CONTAINER)
    try:
        container_client.create_container()
    except Exception:
        pass
    container_client.upload_blob(name=blob_name, data=json_bytes, overwrite=True)
    logger.info(f"Uploaded log to Azure Blob: {GENERATED_CONTAINER}/{blob_name}")
    return f"{GENERATED_CONTAINER}/{blob_name}"

def list_generated_presentations():
    blob_service = BlobServiceClient.from_connection_string(BLOB_CONN)
    container_client = blob_service.get_container_client(GENERATED_CONTAINER)
    try:
        return [b.name for b in container_client.list_blobs()]
    except Exception as e:
        logger.warning(f"Failed to list generated PPTs: {e}")
        return []
