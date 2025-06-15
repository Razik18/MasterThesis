import os
import json
import sys
import logging
import requests
from flask import Blueprint, render_template, request, jsonify
from paddleocr import PaddleOCR
from datetime import datetime
import pandas as pd
from app.config import GLOBAL_STATE

# Determine the path to the settings file relative to if its an executable or not (frozen)
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(__file__)

settings_file = os.path.join(application_path, "settings.json")


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

ocr=Blueprint('ocr', __name__)

def perform_ocr(image_url):
    try:
        response = requests.get(image_url, stream=True)
        if response.status_code == 200:
            ocr = PaddleOCR(use_angle_cls=True, lang='en')
            result = ocr.ocr(response.content, cls=True)
            extracted_data = []
            for line in result[0]:
                text = line[1][0]
                score = line[1][1]
                extracted_data.append({"text": text, "score": score})
            return extracted_data
    except Exception as e:
        logger.error(f"Error performing OCR: {e}")
    return []

def match_images_to_folders(images, folders):
    image_folder_map = {}
    for image in images:
        folder_id = image.get("folderParentId")
        if folder_id is not None:
            folder = next((f for f in folders if f.get("id") == folder_id), None)
            while folder and not folder.get("hasMetadata"):
                folder_id = folder.get("folderParentId")
                folder = next((f for f in folders if f.get("id") == folder_id), None)
            if folder:
                image_folder_map[image["id"]] = folder["id"]
    return image_folder_map

def get_field_names_by_ids(field_ids):
    metadata = GLOBAL_STATE['service'].client.get_metadata_fields()
    fields = metadata.get("fields", [])
    field_names = {field["id"]: field["name"] for field in fields if field["id"] in field_ids}
    return {field_id: field_names.get(field_id) for field_id in field_ids if field_id in field_names}

def update_folder_metadata(folder_id, metadata):
    folder_data = GLOBAL_STATE['service'].client.get_folder(folder_id)
    if folder_data.get("hasMetadata"):
        metadata_values = [{"fieldId": int(field_id), "resourceId": int(folder_id), "content": content}
                           for field_id, content in metadata.items()]
        return GLOBAL_STATE['service'].update_metadata_values(metadata_values)
    return False

def load_settings():
    try:
        with open(settings_file, "r") as file:
            return json.load(file)
    except FileNotFoundError:
        return {"templates": {}}

# Main Routes
@ocr.route('/ocr')
def home():
    return render_template('index.html')

@ocr.route('/rename')
def rename():
    return render_template('rename.html')

@ocr.route('/transfer')
def transfer():
    return render_template('transfer.html')

@ocr.route('/perform_ocr', methods=['POST'])
def perform_ocr_route():
    image_id = request.form.get('image_id')
    image_url = GLOBAL_STATE['service'].get_image_url(image_id)
    macro_url = GLOBAL_STATE['service'].get_macro_url(image_id)
    if image_url:
        extracted_data = perform_ocr(image_url)
        return jsonify({"data": extracted_data, "image_url": image_url, "macro_url": macro_url})
    return jsonify({"error": "Image not found"}), 404


@ocr.route('/update_metadata', methods=['POST'])
def update_metadata():
    data = request.json
    metadata_values = []
    image_id = data.get('image_id')
    metadata = data.get('metadata', {})

    for field_id, content in metadata.items():
        if int(field_id) == 28:  # patient id
            folder_id = GLOBAL_STATE['service'].get_folder_id_for_image(image_id)
            while folder_id:
                if update_folder_metadata(folder_id, {field_id: content}):
                    break
                folder_id = GLOBAL_STATE['service'].get_parent_folder_id(folder_id)
        else:
            metadata_values.append({
                "fieldId": int(field_id),
                "resourceId": int(image_id),
                "content": content
            })

    if metadata_values:
        response = GLOBAL_STATE['service'].update_metadata_values(metadata_values)
        if response is None:
            return jsonify({"error": "Failed to update metadata"}), 500

    return jsonify({"message": "Metadata updated successfully"})


@ocr.route('/settings')
def settings():
    metadata = GLOBAL_STATE['service'].get_metadata()
    return render_template('settings.html', metadata=metadata, templates=settings.get("templates", {}))

@ocr.route('/save_settings', methods=['POST'])
def save_settings():
    global settings
    data = request.get_json()
    template_name = data.get('template_name')
    metadata_map = data.get('metadata_map')
    protocol_stain_map = data.get('protocol_stain_map')
    naming_pattern = data.get('naming_pattern')
    use_patient_id = data.get('use_patient_id')  
    use_image_name = data.get('use_image_name')  

    metadata = json.loads(metadata_map) if metadata_map else {}
    protocols = json.loads(protocol_stain_map) if protocol_stain_map else {}

    if template_name:
        settings["templates"][template_name] = {
            "metadata": {int(k): v for k, v in metadata.items()},
            "protocols": protocols,
            "naming_pattern": naming_pattern,
            "use_patient_id": use_patient_id,  
            "use_image_name": use_image_name  
        }

    with open(settings_file, "w") as file:
        json.dump(settings, file, indent=4)

    return jsonify({"message": "Settings updated successfully"})



@ocr.route('/get_stain_options', methods=['GET'])
def get_stain_options():
    field_id = 31
    stain_options = GLOBAL_STATE['service'].get_metadata_field_options(field_id)
    if stain_options:
        stain_values = {option["id"]: option["value"] for option in stain_options}
        return jsonify(stain_values)
    return jsonify({"error": "Failed to retrieve stain options"}), 500

@ocr.route('/get_templates', methods=['GET'])
def get_templates():
    return jsonify(settings.get("templates", {}))

@ocr.route('/save_template', methods=['POST'])
def save_template():
    template_name = request.form.get('template_name')
    protocol_stain_map = request.form.get('protocol_stain_map')
    if template_name and protocol_stain_map:
        settings['templates'][template_name] = json.loads(protocol_stain_map)
        with open(settings_file, "w") as file:
            json.dump(settings, file, indent=4)
        return jsonify({"message": "Template saved successfully"})
    return jsonify({"error": "Failed to save template"}), 500

@ocr.route('/get_metadata_fields', methods=['GET'])
def get_metadata_fields():
    metadata = GLOBAL_STATE['service'].get_metadata()
    return jsonify(metadata)

#not used anymore
@ocr.route('/get_metadata_for_imageset', methods=['POST'])
def get_metadata_for_imageset_route():
    data = request.get_json()
    imageset_id = data.get('imagesetId')
    if not imageset_id:
        return jsonify({"error": "imagesetId is required"}), 400
    # Fetch metadata for all images in the specified image set
    metadata = GLOBAL_STATE['service'].get_metadata_for_imageset(imageset_id)
    return jsonify(metadata)







@ocr.route('/load_images_with_metadata', methods=['POST'])
def load_images_with_metadata():
    current_settings = load_settings()
    
    source_repository_id = request.form.get('source_repository_id')
    selected_template = request.form.get('selected_template')
    
    # Fetch images and folders
    images = GLOBAL_STATE['service'].get_all_images_by_imageset_id(source_repository_id)
    folders = GLOBAL_STATE['service'].get_folders_by_imageset_id(source_repository_id)
    image_folder_map = match_images_to_folders(images, folders)

    # Fetch metadata for imageset
    metadata = GLOBAL_STATE['service'].get_metadata_for_imageset(source_repository_id)

    # Get metadata fields from the selected template
    try:
        template_settings = current_settings['templates'][selected_template]
    except KeyError:
        return jsonify({"error": "Selected template not found"}), 404

    template_metadata = template_settings['metadata']
    template_metadata_fields = template_metadata.keys()
    field_id_to_name = {str(field_id): field_data['name'] for field_id, field_data in template_metadata.items()}

    # Check if patient id details should be included
    use_patient_id = template_settings.get("use_patient_id", False)

    # Fetch detailed stain options
    stain_field_id = 31  
    stain_options = GLOBAL_STATE['service'].get_metadata_field_options(stain_field_id)
    stain_value_map = {option["id"]: option["value"] for option in stain_options} if stain_options else {}

    image_data = []
    for image in images:
        image_id = image.get("id")
        created_date = image.get("created")
        date = datetime.strptime(created_date, "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d-%m-%Y") if created_date else None
        folder_id = image_folder_map.get(image_id)

        # Extract label and macro URLs
        label_url = None
        macro_url = None
        for associated_image in image.get("associatedImages", []):
            if associated_image["type"] == "label":
                label_url = associated_image["signedURL"]
            elif associated_image["type"] == "macro":
                macro_url = associated_image["signedURL"]

        # Extract metadata for the image
        image_metadata = {
            field_id_to_name[str(item['fieldId'])]:
                (stain_value_map.get(item['content'], item['content']) if str(item['fieldId']) == str(stain_field_id)
                 else str(item['content']))
            for item in metadata
            if item['resourceId'] == image_id and str(item['fieldId']) in template_metadata_fields
        }

        field_values = {
            "Image Name": str(image.get("name")),
            "Date": str(date),
            "Image ID": str(image_id),
            "Label URL": str(label_url),
            "Macro URL": str(macro_url),
            **image_metadata
        }

        # If patient ID is enabled, add folder and patient info
        if use_patient_id:
            patient_id_value = "! Missing"
            if folder_id:
                patient_meta = [item for item in metadata if item['resourceId'] == folder_id and str(item['fieldId']) == "28"]
                if patient_meta:
                    patient_id_value = str(patient_meta[0]['content'])
            field_values["Folder ID"] = str(folder_id) if folder_id is not None else "! Missing"
            field_values["Patient ID"] = patient_id_value

        # Ensure every field from the template is present
        for field_id, field_data in template_metadata.items():
            field_name = field_data['name']
            if field_name not in field_values or not field_values[field_name]:
                field_values[field_name] = "! Missing"

        image_data.append(field_values)

    # Create and process DataFrame
    df = pd.DataFrame(image_data).astype(str)
    df = df.fillna("! Missing").replace("nan", "! Missing").replace("None", "! Missing")

    json_data = df.to_dict(orient='records')
    return jsonify(json_data)


@ocr.route('/load_images_for_transfer', methods=['POST'])
def load_images_for_tranfer():
    source_repository_id = request.form.get('source_repository_id')

    images = GLOBAL_STATE['service'].get_all_images_by_imageset_id(source_repository_id)
    imageset_name = GLOBAL_STATE['service'].get_imageset_name(source_repository_id)

    image_data = []
    for image in images:
        image_id = image.get("id")
        created_date = image.get("created")
        date = (
            datetime.strptime(created_date, "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d-%m-%Y")
            if created_date
            else None
        )
        storagekey = image.get("storageKey")
        field_values = {
            "Image Name": str(image.get("name")),
            "Date": str(date),
            "Image ID": str(image_id),
            "storageKey": str(storagekey),
        }
        image_data.append(field_values)

    df_images = pd.DataFrame(image_data).astype(str)
    df_images = df_images.fillna("! Missing").replace("nan", "! Missing").replace("None", "! Missing")
    json_data = df_images.to_dict(orient='records')

    attachments_response = GLOBAL_STATE['service'].get_attachments(source_repository_id)
    attachment_data = []

    if attachments_response:
        if isinstance(attachments_response, dict) and "attachments" in attachments_response:
            attachments_list = attachments_response["attachments"]
        else:
            attachments_list = attachments_response

        for att in attachments_list:
            if isinstance(att, dict):
                attachment_data.append({
                    "Attachment Name": str(att.get("name", "! Missing")),
                    "storageKey": str(att.get("storageKey", "! Missing"))
                })
            else:
                attachment_data.append({
                    "Attachment Name": str(att),
                    "storageKey": str(att)
                })

    df_attachments = pd.DataFrame(attachment_data).astype(str)
    df_attachments = df_attachments.fillna("! Missing").replace("nan", "! Missing").replace("None", "! Missing")
    attachments_json = df_attachments.to_dict(orient='records')

    return jsonify({
        "images": json_data,
        "imageset_name": imageset_name,
        "attachments": attachments_json
    })



@ocr.route('/update_image_name', methods=['POST'])
def update_image_name_route():
    image_id = request.form.get('image_id')
    new_name = request.form.get('new_name')
    new_name_with_extension = f"{new_name}.svs"
    success = GLOBAL_STATE['service'].update_image_name(image_id, new_name_with_extension)
    if success:
        return jsonify({"message": "Image name updated successfully"})
    return jsonify({"error": "Failed to update image name"}), 500

@ocr.route('/update_image_folder', methods=['POST'])
def update_image_folder_route():
    image_id = request.form.get('image_id')
    repository_id = request.form.get('repository_id')
    folder_name = request.form.get('folder_name')
    is_case = request.form.get('is_case') == 'true'

    folders = GLOBAL_STATE['service'].get_folders_by_imageset_id(repository_id)
    folder_parent_id = None
    for folder in folders:
        if folder.get('label') == folder_name and folder.get('hasMetadata') == is_case:
            folder_parent_id = folder.get('id')
            break

    if folder_parent_id is not None:
        success = GLOBAL_STATE['service'].update_image_folder(image_id, folder_parent_id)
        if success:
            return jsonify({"message": "Image folder updated successfully", "folder_id": folder_parent_id})
        return jsonify({"error": "Failed to update image folder"}), 500

    folder_parent_id = GLOBAL_STATE['service'].create_folder(
        label=folder_name,
        image_set_id=repository_id,
        has_metadata=is_case
    )

    if folder_parent_id is not None:
        success = GLOBAL_STATE['service'].update_image_folder(image_id, folder_parent_id)
        if success:
            return jsonify({"message": "Image folder updated successfully after creating a new folder", "folder_id": folder_parent_id})
        return jsonify({"error": "Failed to update image folder after creating a new folder", "folder_id": folder_parent_id}), 500

    return jsonify({"error": "Failed to create a new folder"}), 500

settings = {}
try:
    with open(settings_file, "r") as file:
        settings = json.load(file)
except FileNotFoundError:
    settings = {"templates": {}}
