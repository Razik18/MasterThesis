from ConcentriqSDK.concentriq_client import ConcentriqAPIClient
import base64
import logging
from flask import session

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ConcentriqService:
    def __init__(self, endpoint=None, email=None, password=None, pac_url = None, logging_level=logging.INFO, logger=None):
        if logger is not None:
            self.logger = logger
        else:
            self.logger = logging.getLogger(__name__)
            self.logger.setLevel(logging_level)
        
        self.pac_url = pac_url
        self.client = ConcentriqAPIClient(endpoint, email, password, logger=self.logger)

    def is_logged(self):
        return self.client._whoami() is not None

    # def authenticate(self, username, password):
    #     session['credentials'] = f"{username}:{password}"
    #     self.client.session.headers.update({'Authorization': f"Basic {base64.b64encode(f'{username}:{password}'.encode()).decode('utf-8')}"})
    #     return True

    def get_image_url(self, image_id):
        image_data = self.client.get_image(image_id)
        if image_data:
            associated_images = image_data.get("associatedImages", [])
            for image in associated_images:
                if image["type"] == "label":
                    return image["signedURL"]
        return None

    def get_macro_url(self, image_id):
        image_data = self.client.get_image(image_id)
        if image_data:
            associated_images = image_data.get("associatedImages", [])
            for image in associated_images:
                if image["type"] == "macro":
                    return image["signedURL"]
        return None

    def get_metadata(self):
        filters = {"resourceType": ["image"]}
        metadata = self.client.get_metadata_fields(filters)
        if metadata:
            fields = metadata.get("fields", [])
            return [{"name": field["name"], "id": field["id"]} for field in fields if field.get("resourceType") == "image"]
        else:
            logger.error("Failed to retrieve metadata fields.")
            return []

    def get_metadata_field_options(self, field_id):
        filters = {"fieldId": [field_id]}
        metadata = self.client.get_metadata_fields(filters)
        if metadata:
            fields = metadata.get("fields", [])
            if fields:
                field = fields[0]
                return field.get("dropdownOptions", [])
        logger.error(f"Failed to retrieve metadata field options for field ID {field_id}.")
        return []

    def get_all_images_by_imageset_id(self, imageset_id):
        filters = {"imageSetId": [imageset_id]}
        pagination = {"rowsPerPage": 100, "page": 1}
        all_images = []
        while True:
            images_data = self.client.get_images(filters, pagination)
            if images_data:
                images = images_data.get("images", [])
                if not images:
                    break
                all_images.extend(images)
                pagination["page"] += 1
            else:
                logger.error(f"Failed to retrieve images for Repository ID {imageset_id}")
                break
        return all_images

    def get_metadata_for_imageset(self, imageset_id):
        all_metadata = []
        filters = {"imageSetId": [imageset_id]}
        metadata = self.client.get_metadata_values(filters)
        if metadata:
            all_metadata.extend(metadata)
        return all_metadata


    def get_folders_by_imageset_id(self, imageset_id):
        filters = {"imageSetId": [imageset_id]}
        pagination = {"rowsPerPage": 100, "page": 1}
        all_folders = []
        while True:
            folders_data = self.client.get_folders(filters=filters, pagination=pagination)
            if folders_data:
                folders = folders_data.get("folders", [])
                all_folders.extend(folders)
                if len(folders) < pagination["rowsPerPage"]:
                    break
                pagination["page"] += 1
            else:
                logger.error(f"Failed to retrieve folders for Repository ID {imageset_id}")
                break
        return all_folders
    
    def get_folders_with_attachments_by_imageset_id(self, imageset_id):
        filters = {"imageSetId": [imageset_id],"hasAttachments":True}
        pagination = {"rowsPerPage": 100, "page": 1}
        all_folders = []
        while True:
            folders_data = self.client.get_folders(filters=filters, pagination=pagination)
            if folders_data:
                folders = folders_data.get("folders", [])
                all_folders.extend(folders)
                if len(folders) < pagination["rowsPerPage"]:
                    break
                pagination["page"] += 1
            else:
                logger.error(f"Failed to retrieve folders for Repository ID {imageset_id}")
                break
        return all_folders
    
    def get_attachments(self, imageset_id):
        folders = self.get_folders_with_attachments_by_imageset_id(imageset_id)

        folder_ids = []
        for folder in folders:
            folder_ids.append(folder["id"])

        if not folder_ids:
            self.logger.warning(
                f"No folders with attachments found for imageset id {imageset_id}"
            )
            return None

        filters = {"folderId": folder_ids}
        attachments = self.client.get_attachment(filters)
        return attachments if attachments else None

    def get_metadata_for_folders(self, folder_ids):
        all_metadata = []
        for i in range(0, len(folder_ids), 500):
            batch_ids = folder_ids[i:i+500]
            filters = {"folderId": batch_ids}
            metadata = self.client.get_metadata_values(filters)
            if metadata:
                all_metadata.extend(metadata)
        return all_metadata

    def get_folder_id_for_image(self, image_id):
        image_data = self.client.get_image(image_id)
        return image_data.get("folderParentId") if image_data else None
    
    def get_imageset_name(self, imageset_id):
        imageset_data = self.client.get_imageset(imageset_id)
        return imageset_data.get("name") if imageset_data else None
 
    def get_parent_folder_id(self, folder_id):
        folder_data = self.client.get_folder(folder_id)
        return folder_data.get("folderParentId") if folder_data else None

    def update_metadata_values(self, metadata_values):
        return self.client.update_metadata_values(metadata_values)
    
    def update_image_name(self, image_id, new_name):
        data = {"name": new_name}
        response = self.client.update_image(image_id, data)
        if response:
            logger.info(f"Image name updated successfully for image_id: {image_id}")
            return True
        logger.error(f"Failed to update image name for image_id: {image_id}")
        return False
    
    def update_image_folder(self, image_id, folder_parent_id):
        """Updates the folderParentId of an image."""
        data = {"folderParentId": folder_parent_id}
        response = self.client.update_image(image_id, data)
        if response:
            logger.info(f"Image folder updated successfully for image_id: {image_id}")
            return True
        logger.error(f"Failed to update image folder for image_id: {image_id}")
        return False
    
    def update_image_imageset(self, image_id, imageset_id):
        """Updates the imageSetIdId of an image."""
        data = {"imageSetId": imageset_id}
        response = self.client.update_image(image_id, data)
        if response:
            logger.info(f"Repository updated successfully for image_id: {image_id}")
            return True
        logger.error(f"Failed to update repository for image_id: {image_id}")
        return False
    
    def create_folder(self, label, image_set_id, has_metadata, folder_parent_id=None):
        """Creates a new folder."""
        folder_parent_id = folder_parent_id if folder_parent_id is not None else -1
        response = self.client.create_folder(
            label=label,
            folderParentId=folder_parent_id,
            imageSetId=image_set_id,
            hasMetadata=has_metadata,
        )
        if response and 'id' in response:
            logger.info(f"Folder created successfully with ID: {response['id']}")
            return response['id']
        logger.error("Failed to create folder.")
        return None

