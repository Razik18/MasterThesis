# API documentation: https://labcorp-dev.concentriq.proscia.com/api-documentation
"""
Example for filter.fields
"fields": {
    "<metadata-field-id>": {
    "contentType": "<data type>",
    "resourceType": "<field type>",
    "values": [
        <your value query here>
        ]
    }
}
"""
import requests
import base64
import json
import re
import logging
from io import StringIO

class ConcentriqAPIClient:
    def __init__(self, endpoint:str=None, email:str=None, password:str=None, logger=None,user_organization_id:int=1):
        if logger:
            self.logger = logger
        else:
            self.logger = logging.getLogger(__name__)
            self.logger.setLevel(logging.INFO)

        self.endpoint = endpoint.rstrip('/') + '/api' if '/api' not in endpoint else endpoint.rstrip('/')
        self.session = requests.Session()
        auth_header = 'Basic {}'.format(base64.b64encode(bytes('{}:{}'.format(email, password), 'utf-8')).decode('utf-8'))
        self.session.headers.update({'Authorization': auth_header})
        
        self.user_organization_id = user_organization_id

        self.metadata_fields = None

        self.config = self._get_instance_config()
        self.logger.debug('config', self.config)
        if self.config:
            self.aws_access_key_id = self.config['storageSystems']['1']['AWS']['accessKeyId']
            self.aws_region = self.config['storageSystems']['1']['AWS']['region']
            self.s3_bucket_name = self.config['storageSystems']['1']['AWS']['S3']['bucket']
            if 'endpointExternal' in self.config['storageSystems']['1']['AWS']['S3']:
                self.s3_endpoint = self.config['storageSystems']['1']['AWS']['S3']['endpointExternal'] 
            else:
                self.s3_endpoint = f'https://s3.{self.aws_region}.amazonaws.com' 

        self.logger.info('Logging in to {}'.format(self.endpoint))
        
    def _handle_response(self, response):
        if not response.ok:
            self.logger.error(f"Request failed with status {response.status_code}: {response.text}")
            return None
        
        try:
            response.raise_for_status()
        except Exception as e:
            self.logger.error(f"Request failed with status {response.status_code}: {response.text}")
            self.logger.exception(e)
            return None
    
        if response.status_code == 204:
            self.logger.debug("No content to return (status 204).")
            return None
        
        self.logger.debug(f"Successful API response: {response.status_code}")
        self.logger.debug(f"Response content: {response.text}")
        return response
    
    def _handle_and_parse_response_text(self, response): 
        response = self._handle_response(response)
        if response:
            return response.text
        return None
    
    def _handle_and_parse_response_text_json(self, response): 
        response = self._handle_response(response)
        if response:
            return json.loads(response.text)
        return None
        
    def _handle_and_parse_response_text_data(self, response): 
        response = self._handle_response(response)
        if response:
            return json.loads(response.text).get('data', [])
        return None
    
    def _handle_and_parse_response_headers(self, response):
        response = self._handle_response(response)
        if response:
            return response.headers
        return None

    def _get_instance_config(self):
        response = self.session.get('{}/config'.format(self.endpoint[:-4]))
        if response is not None:
            match = re.search(r'.+JSON.parse\(\'(.+)\'\)\;', response.text)
            if match:
                return json.loads(match.groups()[0])
        return None
    
    def _whoami(self):
        response = self.session.get('{}/v3/auth/whoami'.format(self.endpoint))
        return self._handle_and_parse_response_text_json(response)

    ### Authentication
    def get_sign_s3_url(self, resourceType:str, resourceId:int, payload:str, nonce:str, canonicalRequest:str):
        """Signs an S3 URL for uploading data with multipart.
        Args:
            resourceType (str): The type of resource the storageKey is associated with. 'analysis', 'image', ...
            resourceId (int): The ID of resource the storageKey is associated with.
            payload (str): Data to be signed.
            nonce (str): DateTime to be used as an argument in signature generation.
            canonicalRequest (str): the canonical request that was used to generate the string to sign. Used to validate the string to sign.
        Returns:
            signature (str): Computed signature
        """
        response = self.session.get(
            '{}/auth/sign/s3-multipart-url/{}/{}'.format(self.endpoint, resourceType, resourceId),
            params={
                'payload':payload,
                'nonce':nonce,
                'canonicalRequest':canonicalRequest
            }
        )
        return self._handle_and_parse_response_text_data(response)

    ### Analysis
    def delete_analysis_job(self, job_id:int):
        """Deletes specified job.
        Args:
            job_id (int): ID of the analysis job to delete
        """
        response = self.session.delete('{}/analysis/jobs/{}'.format(self.endpoint, job_id))
        return self._handle_and_parse_response_text_data(response)
    
    def get_analysis_job_markups(self, job_id:int):
        """Returns urls of markups and constituent objects
        Args:
            job_id (int): ID of the analysis job get markups for
        Returns:
            chunkURLs (Object): A map keyed by chunk name of signed URLs for each chunk
            indexURL (str): A signed URL for the index file, which holds the structure of the markups
        """
        response = self.session.get('{}/analysis/jobs/{}/markups'.format(self.endpoint, job_id))
        return self._handle_and_parse_response_text_data(response)
    
    # In API documentation, there is 'markups' instead of 'overlays'
    def get_analysis_job_overlays(self, job_id:int):
        """Returns information about raster image overlays for an Analysis Job
        Args:
            job_id (int): ID of the analysis job get markups for
        Returns:
            chunkURLs (Object): A map keyed by chunk name of signed URLs for each chunk
            indexURL (str): A signed URL for the index file, which holds the structure of the markups
        """
        response = self.session.get('{}/analysis/jobs/{}/overlays'.format(self.endpoint, job_id))
        return self._handle_and_parse_response_text_data(response)
    
    def get_analysis_jobs(self, filters:dict={}):
        """Returns all jobs with specified filters.
        Args:
            filters (dict, optional): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR). Defaults to {}.
                imageSetId (list[int]): An array of imageSetIds to filter by the image set that jobs were run within.
                imageId (list[int]): An array of image ids for images on which analysis jobs were run.
                analysisJobId (list[int]): An array of ids of specific analysis jobs to match.
                fields (dict): An map with the metadataFieldId as the key and the value matching the contentType.
        Returns:
            jobs (list[Object]):
                id (int): The ID of the Job
                label (str): The name of the job
                imageId (int): Id of the image this job is associated with
                moduleId (int): Id of the module this job is associated with
                size (float): Number of chunks which compose the markups
                startTime (Date): Time that the job was started
                finishTime (Date): Time that the job was finished
                errorMessage (str): Failure message for the job
                failed (bool): If the job has failed
                overlays (any): ??
                userId (int): ID of the user who created the job
                boundsString (str): bounds of the entire job
                colorMap (dict): Properties about job
                sharePermissions (dict): Permissions inhereted from the imageSet
                hasMarkups (dict): Determines if this job has markups associated with it
        """
        response = self.session.get(
            '{}/analysis/jobs'.format(self.endpoint),
            params={'filters' : json.dumps(filters)}
        )
        return self._handle_and_parse_response_text_data(response)
    
    def update_analysis_job(self, job_id:int, completed:bool, overlays:dict=None, hasMarkups:bool = None):
        """
        Updates an Analysis Job.
        This function will make unable markups. So markups need to be added after that this function is called. If overlays is None, it will also remove all overlays, so need to combine with another function to keep overlays. 
        So, basically this function need to be called before adding markup.
        Also, on one image overlay and markup cannot be set in same time for now.
        Args:
            job_id (int): ID of the Analysis Job
            completed (bool): status of the Analysis Job.
            If True, it allows to set (or update if already set) 'finishTime' in analysis job data, and add 'Date received' and 'Time received' on Concentriq.
            hasMarkups (bool, optional): whether the Analysis Job has vector/object markups.
            True if show on Concentriq the markup button (could be placed even if the analysis have no markup), but will show error when click on button after calling this function. False will remove the button. In both case, it will remove 'markupsStorageKey' from analysis job values.
            In API documentation, it is as non-optional parameters but apparently it is not.
            overlays (dict): object describing the raster overlays for this Analysis Job.
            If this is not specifiy, this function will remove overlays. So to keep already put overlay, need to get in with get_analysis_job_overlays. 
            If completed is not True, it will not add overlay.
        """
        data = {'completed': completed}
        if overlays:
            data['overlays'] = overlays
        if hasMarkups:
            data['hasMarkups'] = hasMarkups
        response = self.session.patch(
            '{}/analysis/jobs/{}'.format(self.endpoint, job_id),
            json=data
        )
        return self._handle_and_parse_response_text_data(response)
    
    def create_analysis_job(self, imageId:int, moduleId:int, label:str):
        """ Creates an Analysis Job.
        Args:
            imageId (int): ID of the Image to which the Analysis Job will be associated
            moduleId (int): ID of the Analysis Module associated with this Analysis Job
            label (str): text string describing the Analysis Job
        Returns:
            dict: dictionary with created 'id' (int)
        """
        response = self.session.post(
            '{}/analysis/jobs'.format(self.endpoint),
            json={
                'imageId': imageId,
                'moduleId': moduleId,
                'label': label
            }
        )
        return self._handle_and_parse_response_text_data(response)
    
    ### Annotation
    def delete_annotation(self, annotation_id:int):
        """Deletes the annotation, removing it from it's associated image.
        Args:
            annotation_id (int): Annotation unique ID.
        Returns:
            success (str): Message detailing the successful deletion of the annotation.
        """
        response = self.session.delete('{}/annotations/{}'.format(self.endpoint, annotation_id))
        return self._handle_and_parse_response_text_data(response)
    
    def get_annotation(self, annotation_id:int):
        """Returns an annotation by ID
        Args:
            annotation_id (int): Annotation unique ID.
        Returns:
            id (int): Unique id of the annotation.
            text (str): A text provided upon creation of the annotation.
            shape (str): The type of shape of the annotation.
            shapeString (str): The compositional components of the shape.
            boundsString (str): Represents a rectangle that bounds the shape on the image.
            captureBounds (str): Represents the viewing image the annotation was created in.
            color (str). Color of the annotation represented in a hexidecimal string.
            size (int): Area of the annotation
            isSegmenting (bool): Determines if the annotation is used for segmentation
            isNegative (bool): Specifies whether this annotation represents a negative region
            imageId (int): The ID of the image this annotation is associated with.
            userId (int): The ID of the user who created the annotation.
            labelOrderX (str): The x order for the annotation to be labeled by in the viewer.
            labelOrderY (str): The y order for the annotation to be labeled by in the viewer.
            creatorName (str): The name associated with userId
            created (date): Date the annotation was created
        """
        response = self.session.get('{}/annotations/{}'.format(self.endpoint, annotation_id))
        return self._handle_and_parse_response_text_data(response)
    
    def get_annotations(self, filters:dict):
        """Returns all annotations the authenticated user has access to, with optional search query parameters
        TODO
        Args:
            filters (dict): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR).
                annotationId (list[int]): An array of annotationIds to filter by.
                imageId (list[int]): An array of imageIds to filter by.
                text (list[str]): An array of annotation text to filter by.
                fields (Object): An object with the metadataFieldId as the key and the value matching the contentType.
        Returns:
            annotations (list[dict])
                id (int): Number Unique id of the annotation.
                text (str): A text provided upon creation of the annotation.
                shape (str): The type of shape of the annotation.
                shapeString (str): The compositional components of the shape.
                boundsString (str): Represents a rectangle that bounds the shape on the image.
                captureBounds (str):Represents the viewing image the annotation was created in.
                color (str): Color of the annotation represented in a hexidecimal string.
                size (int): Area of the annotation.
                isSegmenting (bool): Determines if the annotation is used for segmentation.
                isNegative (bool): Specifies whether this annotation represents a negative region.
                imageId (int): The ID of the image this annotation is associated with.
                userId (int): The ID of the user who created the annotation. 
                labelOrderX (str): The x order for the annotation to be labeled by in the viewer.
                labelOrderY (str): The y order for the annotation to be labeled by in the viewer.
                creatorName (str): The name associated with userId.
                created (date): Date the annotation was created.
        """
        response = self.session.get(
            '{}/annotations'.format(self.endpoint),
            params={'filters':json.dumps(filters)}
        )
        return self._handle_and_parse_response_text_data(response)
    
    def update_annotation(self, annotation_id:int, data:dict):
        """Updates properties of an annotation.
        Args:
            annotation_id (int): ID of the annotation.
            data (dict): _description_
        Returns:
            text (str): The text of the annotation.
            isSegmenting (bool): whether the annotation is a segmentation.
            imageId (int): the image to associate the annotation with.
            captureBounds (str): A string defining the bounds of the viewport when the annotation was created.
            shapeString (str): A string defining the shape of the annotation.
            labelOrderX (str): The string value for 'x' in the (x, y) coordinate labeling of the annotation.
            labelOrderY (str): The string value for 'y' in the (x, y) coordinate labeling of the annotation.
            color (str): the color of the annotation.
            isNegative (bool): Specifies whether this annotation represents a negative region.
        """
        response = self.session.patch(
            '{}/annotations/{}'.format(self.endpoint, annotation_id),
            json=data)
        return self._handle_and_parse_response_text_data(response)
    
    def create_annotation(self, annotation:dict):
        """Creates an annotation.
        Args:
            annotation (dict):
                annotationClassId (int): only from V4.2
                imageId (int): The image to associate the new annotation with.
                color (str): the color of the annotation.
                shape (str): The type of shape of the annotation. 'free'
                shapeString (str): A string defining the shape of the annotation.
                isNegative (bool): Specifies whether this annotation represents a negative region.
                text (str): The text of the annotation.
                captureBounds (str): A string defining the bounds of the viewport when the annotation was created. None is ok.
                isSegmenting (bool): whether the annotation is a segmentation. False is ok.
                bounds (str): A string defining the bounds of the annotation. None is ok.
        Returns:
            id: annotation id.
        """
        response = self.session.post(
            '{}/annotations'.format(self.endpoint),
            json=annotation)
        return self._handle_and_parse_response_text_data(response)
    
    ### Attachments
    def delete_attachment(self, attachment_id:int):
        """Deletes an attachment.
        Args:
            attachment_id (int): ID of the attachment to delete.
        """
        response = self.session.delete('{}/attachments/{}'.format(self.endpoint, attachment_id))
        return self._handle_and_parse_response_text_data(response)
    
    def get_attachment(self, filters:dict):
        """Gets all attachments based on the provided filter.
        Args:
            filters (dict): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR). 
                folderId (list[int]): Filters attachments by folderId.
                attachmentId (list[int]): Filters attachments by uid. 
                storageKey (list[str]): Filters attachments by storageKey.
                imageSetId (list[int]): Filters attachments by imageSetId.
        Returns:
            attachments (list[dict]): Array of attachments.
                id (int): ID of the attachment.
                name (str): Filename of the attachment.
                storageKey (str): Location of the file.
                uploaderId (int): ID of the user who uploaded the attachment.
                resourceType (str): Type of resource the attachment is associated with.
                resourceId (int): ID of the resource the attachment is associated with.
        """
        response = self.session.get(
            '{}/attachments'.format(self.endpoint),
            params={'filters':json.dumps(filters)}
        )
        return self._handle_and_parse_response_text_data(response)
    
    def create_attachment(self, name:str, resourceType:str, resourceId:int):
        """Creates an attachment and returns a signing key for upload.
        Args:
            name (str): The name of the file.
            resourceType (str): Type resource to associate the file with. Only 'folder' works? 'image' not works.
            resourceId (int): ID of the resource to associate the file with.
        Returns:
            attachmentId (int): The ID of the newly created field.
            storageKey (str): The storageKey of the newly created file. (Used to upload file on this storagekey?)
        """
        response = self.session.post(
            '{}/attachments'.format(self.endpoint),
            json={
                'name': name,
                'resourceType': resourceType,
                'resourceId': resourceId
            })
        return self._handle_and_parse_response_text_data(response)
    
    ### Folder
    def delete_folder(self, folder_id:int):
        """Deletes the specified folder.
        Args:
            folder_id (int): ID of the folder to be deleted.
        Returns:
            success (str): Message detailing the successful deletion of the folder.
        """
        response = self.session.delete('{}/folders/{}'.format(self.endpoint, folder_id))
        return self._handle_and_parse_response_text_data(response)
    
    def get_folder(self, folder_id:int):
        """Returns a folder the authenticated user has access to.
        Args:
            folder_id (int): Unique ID of the folder.
        Returns:
            id (int): Unique ID of the folder.
            label (str): Label of the folder.
            imageSetId (int): ID of the image set to which this folder belongs.
            folderParentId (int): ID of the folder that this folder is within. can be Null for root level folders.
            imageSetName (str): Name of the image set to which this folder belongs.
            hasMetadata (bool): Whether this folder is a case.
            folders.hasAttachments (bool): Whether this folder has attachments associated
            rank (int): The order of the file in the navigation structure
            ownerId (int): ID of the owner of the image set to which this folder belongs.
            sharePermissions (dict): Permissions
        """
        response = self.session.get('{}/folders/{}'.format(self.endpoint, folder_id))
        return self._handle_and_parse_response_text_data(response)
    
    # TODO check for hasReasonForChange if works like this or need to pass as other format
    def get_folders(self, hasReasonForChange:bool=None, includeMetadata:bool=None, idsOnly:bool=None, filters:dict={}, pagination:dict={}):
        """Returns all folders the authenticated user has access to, with optional search query parameters.
        Args:
            hasReasonForChange (bool, optional): Specifies if the folder's imageSet has change control enabled, requiring a reason for change for auditable events. Defaults to None.
            includeMetadata (bool, optional): Optionally adds metadata values to the results. Defaults to None.
            idsOnly (bool, optional): Specifies if the endpoint should return an array of objects containing just the id of the resource. Defaults to None.
            filters (dict, optional): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR). Defaults to {}.
                imageSetId (list[int]): An array of imageSetIds to filter by the image set that this folder belongs to.
                folderId (list[int]): An array of folderIds to filter by.
                hasMetadata (bool): Specifies if a folder is a case or not.
                hasAttachments (bool): Specifies if a folder has attachments. 
                name (list[str]): An array of names to filter by image set name. 
                generalSearch (list[str]): An array of search phrases to search over various fields and case properties. 
                fields (dict): An map with the metadataFieldId as the key and the value matching the contentType. Examles can be found above. 
                image (dict): An objects as defined by the 'images' endpoint filters parameter. Allows filtering by properties of the images to which the folders may belong.
            pagination (dict, optional): Options to specify sort order and size of subset for all image sets to return. Defaults to {}.
                rowsPerPage (int): How many rows to return.
                page (int): Specifies an offset to advance the subset through the set of all image sets.
                sortBy (list[str]): Which property of the case object to sort the entire set by. Options are ['name'], ['created'], ['lastModified'], ['size']. Also accepts a non-array String.
                descending (bool): Specifies which order the values are sorted into.
        Returns:
            folders (list[dict]): Containing object for the folder objects. The following properties are of the objects in the array, not the array itself.
                id (int): Unique ID of the folder.
                label (str): Label of the folder.
                imageSetId (int): ID of the image set to which this folder belongs.
                folderParentId (int): ID of the folder that this folder is within. can be Null for root level folders.
                imageSetName (str): Name of the image set to which this folder belongs. 
                hasMetadata (bool): Whether this folder is a case 
                hasAttachments (bool): Whether this folder has attachments associated 
                rank (int): The order of the file in the navigation structure 
                ownerId (int): ID of the owner of the image set to which this folder belongs. 
                sharePermissions (dict): Permissions granted to the user who called this endpoint for this resource. 
                metadata (dict): Metadata keyed by the fieldId with the content as the value (if the 'includeMetadata' queryParameter is provided).
        """
        params = {
            'filters' : json.dumps(filters),
            'pagination': json.dumps(pagination)
        }
        if hasReasonForChange:
            params['hasReasonForChange'] = hasReasonForChange
        if includeMetadata:
            params['includeMetadata'] = includeMetadata
        if idsOnly:
            params['idsOnly'] = idsOnly
        response = self.session.get(
            '{}/folders'.format(self.endpoint),
            params=params
        )
        return self._handle_and_parse_response_text_data(response)
    
    def update_folder(self, folder_id:int, name:str=None, rank:int=None, imageSetId:str=None):
        """Updates a folder (all argument can be empty, success update with no update).
        Args:
            folder_id (int): folder unique ID.
            name (str, optional): Label/name of the folder. Defaults to None.
            rank (int, optional): The list order of the folder. Defaults to None.
            imageSetId (str, optional): ID of the image set to move this folder to. Defaults to None.
        Returns:
            success (str): Message detailing the successful update of the folder.
        """
        data = {}
        if name:
            data['name']=name
        if rank:
            data['rank']=rank
        if imageSetId:
            data['imageSetId']=imageSetId
        response = self.session.patch(
            '{}/folders/{}'.format(self.endpoint, folder_id),
            json=data
        )
        return self._handle_and_parse_response_text_data(response)
    
    def create_folder(self, label:str, folderParentId:int, imageSetId:str, hasMetadata:bool, cloneSourceId:str = None):
        """Creates a folder.
        Args:
            label (str): Label/name of the folder.
            folderParentId (int): the id of the folder this folder is nested inside. null or -1 for the base directory.
            imageSetId (str): ID of the image set to move this folder to.
            hasMetadata (bool): Whether the folder is a case. True if case, False if folder
            cloneSourceId (str, optional): ID of the source folder to clone.
        Returns:
            Object:
                id (int): folder_id
                success (str): Message detailing the successful update of the folder. Data -> id is the FolderId of the created folder.
        """
        data = {
            'label':label,
            'folderParentId':folderParentId,
            'imageSetId':imageSetId,
            'hasMetadata':hasMetadata,
        }
        if cloneSourceId:
            data['cloneSourceId'] = cloneSourceId
        response = self.session.post(
            '{}/folders'.format(self.endpoint),
            json=data
        )
        return self._handle_and_parse_response_text_data(response)
    
    ### Image
    def delete_image(self, image_id:int):
        """Deletes the specified image.
        Args:
            image_id (int): ID of the image to be deleted.
        Returns:
            success (str): Message detailing the successful deletion of the image.
        """
        response = self.session.delete('{}/images/{}'.format(self.endpoint, image_id))
        return self._handle_and_parse_response_text_data(response)
    
    def get_download_image(self, image_id:int):
        """Responds with a redirect for a link to download the image directly.
        Args:
            image_id (int): image unique ID.
        """
        response = self.session.get('{}/images/{}/download'.format(self.endpoint, image_id), allow_redirects=False)
        return self._handle_and_parse_response_headers(response)
    
    def get_image(self, image_id:int):
        """Returns the requested image data.
        Args:
            image_id (int): image unique ID.

        Returns:
            id (int): Unique id of the image.
            name (str): Label of the image.
            imageSetId (int): Id of the image set to which this image belongs.
            imageSetName (str): Name of the image set to which this image belongs.
            folderParentId (int): Id folder which this image is in, or null.
            ownerId (int): Id of the owner of the image set to which this image belongs. 
            rank (int): Sort order of this image.
            hasMacro (bool): Whether or not this image has an associated macro image.
            hasLabel (bool): Whether or not this image has a label.
            hasOverlays (bool): Whether or not this image has overlays associated.
            hasMultipleZLayers (bool): Whether or not this image has multiple z-layers (z-stack). 
            hasAnnotations (bool): Whether or not this image has annotations associated. 
            hasAnalysisResults (bool): Whether or not this image has analysis results associated.
            rotation (int): The last position of rotation for this image 
            mppx (int): Image resolution in X dimension. 
            mppy (int): Image resolution in Y dimension. 
            imgWidth (int): Image width in pixels. 
            imgHeight (int): Image height in pixels. 
            fileSize (int): size of the image file in bytes. 
            objectivePower (int): Objective power of the image capture - otherwise the image max depth. 
            slideName (str): Name of the underlying slide object. 
            status (int): ID of the ingestion status of this image.
            created (date): Date and time the image was created. 
            storageKey (str): Location of the original image. 
            associatedKey (str): Location of associated images and data produced by the ingestion process. 
            thumbURL (str): URL of the thumbnail for this image. 
            associatedImages (list[dict]): Other images associated with this image. 
                type (str): Type of associated image associatedImages.
                signedURL (str): URL of the associated image. 
            imageData (dict): Data for structuring the image in the viewer imageData.
                imageSources (dict): The sources of the image which contain the actual image data 
                fluorescenceChannels (dict): Data about the fluorescence properties of the image. Null if the image is not fluorescence.
                metadataUrl (dict): Obsolete, This Signed URL was once used for other data about the image. 
            sharePermissions (dict): Permissions granted to the user who called this endpoint for this resource.
        """
        response = self.session.get('{}/images/{}'.format(self.endpoint, image_id))
        return self._handle_and_parse_response_text_data(response)

    # TODO check if really works
    def get_image_annotations_mld(self, image_id:int):
        """Sends an MLD file which is a representation of the annotations for a given image. Note, this is only enabled if the visiopharm integration is enabled.
        Args:
            image_id (int): image unique ID.
        """
        response = self.session.get('{}/images/{}/annotations/mld'.format(self.endpoint, image_id))
        return self._handle_and_parse_response_text_data(response)

    # TODO check if really works
    def get_image_annotations_xml(self, image_id:int):
        """Sends an XML file which is a representation of the annotations for a given image.
        Args:
            image_id (int): image unique ID.
        """
        response = self.session.get('{}/images/{}/annotations/xml'.format(self.endpoint, image_id))
        return self._handle_and_parse_response_text_data(response)

    def get_images(self, filters:dict={}, pagination:dict={}, hasReasonForChange:bool=None, includeMetadata:bool=None, idsOnly:bool=None):
        """Returns all images the authenticated user has access to, with optional search query parameters.
        Args:
            idsOnly (bool): Specifies if the endpoint should return an array of objects containing just the id of the resource.
            includeMetadata (bool): Optionally adds metadata values to the results.
            hasReasonForChange (bool): Specifies if the image's imageSet has change control enabled, requiring a reason for change for auditable events.
            filters (dict, optional): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR). Defaults to {}.
                imageSetId (list[int]): An array of imageSetIds to filter by the image set that this image belongs to.
                imageId (list[int]): An array of imageIds to filter by.
                name (list[str]): An array of names to filter by image name.
                generalSearch (list[str]): An array of search phrases to search over various fields and case properties.
                fields (dict): An object with the metadataFieldId as the key and the value matching the contentType.
                folder (dict): An object as defined by the 'folders' endpoint filters parameter. Allows filtering by properties of the folder an image may belong to.
                analysis (dict): An object as defined by the 'analysis/jobs' endpoint filters parameter. Allows filtering by properties of the analysis job an image may belong to.
                hasOverlays (bool): Whether or not this image has overlays associated.
                hasMultipleZLayers (bool): Whether or not this image has multiple z-layers.
                hasAnnotations (bool): Whether or not this image has annotations associated.
                hasAnalysisResults (bool): Whether or not this image has analysis results associated.
                created (list[dict]): Filters results by the date the image was created
                    start (str): Start date for the query in ISO format 
                    end (str): End date for the query in ISO format
                objectivePower (Object[][]): A 2D array of objects with an operator and operand to filter by objective power.
                fileSize (int):
            pagination (dict, optional): Options to specify sort order and size of subset for all image sets to return. Defaults to {}.
                rowsPerPage (int): How many rows to return.
                page (int): Specifies an offset to advance the subset through the set of all image sets.
                sortBy (list[str]): Which property of the case object to sort the entire set by. Options are ['name'], ['created'], ['lastModified'], ['size']. Also accepts a non-array String.
                descending (bool): Specifies which order the values are sorted into
    
    def import_markups_from_xml(self, analysis_job_id, annotations_xml):
        response = self.session.post(
            '{}/analysis/jobs/{}/markups/importFromXML'.format(self.endpoint, analysis_job_id),
            files={'files': StringIO(annotations_xml)}
        )
        return self._handle_and_parse_response(response)

        Returns:
            images (dict):
                id (int): Unique id of the image.
                name (str): Label of the image.
                imageSetId (int): Id of the image set to which this image belongs.
                imageSetName (str): Name of the image set to which this image belongs.
                folderParentId (int): Id folder which this image is in, or null.
                folderName (str):  Name folder which this image is in, or null.
                ownerId (int): Id of the owner of the image set to which this image belongs.
                rank (int): Sort order of this image.
                hasMacro (bool): Whether or not this image has an associated macro image.
                hasLabel (str): Whether or not this image has a label.
                hasOverlays (str): Whether or not this image has overlays associated.
                hasMultipleZLayers (str): Whether or not this image has multiple z-layers (z-stack).
                hasAnnotations (str): Whether or not this image has annotations associated.
                hasAnalysisResults (str): Whether or not this image has analysis results associated.
                mppx (int): Image resolution in X dimension.
                mppy (int): Image resolution in Y dimension.
                imgWidth (int): Image width in pixels.
                imgHeight (int): Image height in pixels.
                fileSize (int): size of the image file in bytes.
                objectivePower (int): Objective power of the image capture - otherwise the image max depth.
                slideName (str): Name of the underlying slide object.
                status (int): ID of the ingestion status of this image.
                created	Date Date and time the image was created.
                storageKey (str): Location of the original image.
                associatedKey (str): Location of associated images and data produced by the ingestion process.
                thumbURL (str): URL of the thumbnail for this image.
                associatedImages (list[dict]): Other images associated with this image.
                    type (str): Type of associated image
                    signedURL (str): URL of the associated image.
                sharePermissions (dict): Permissions granted to the user who called this endpoint for this resource.
        """
        params = {
            'filters' : json.dumps(filters),
            'pagination': json.dumps(pagination)
        }
        if hasReasonForChange:
            params['hasReasonForChange'] = hasReasonForChange
        if includeMetadata:
            params['includeMetadata'] = includeMetadata
        if idsOnly:
            params['idsOnly'] = idsOnly
        response = self.session.get(
            '{}/images'.format(self.endpoint),
            params=params
        )
        return self._handle_and_parse_response_text_data(response)
    
    def update_image(self, imageId:int, data:dict):
        """Updates image.
        Args:
            imageId (int): image unique ID.
            data (dict):
                name (str): Label/name of the image.
                status (str): the error/uploading/optimizing status of the image. -1 is error. 0 is uploading. 1 is optimizing, and setting to one triggers re-optimization. 2 is success.
                imageSetId (int): ID of the image set to move this image to.
                storageKey (str): new path of the original image file in S3/Filesystem. Admin only feature.
                storageSystemId (int): ID of storage system in which the storageKey will be updated. If none specified, default is used.
                folderParentId (str): the folder to move this image to 
                rank (int): the list order of the image to move to. Decimal values are allowed to specify positioning. 1 is the first item in order. An insert at .5 will set the image as the first item in the navigation order. This should not be used in conjunction with "MoveAfter" properties.
                clearMetadata (bool): Clears any cloned metadata from the duplicated image.
                rotation (int): the rotation of the image
        Returns:
            success (str): Message detailing the successful update of the image.
        """
        response = self.session.patch(
            '{}/images/{}'.format(self.endpoint, imageId),
            json=data
        )
        return self._handle_and_parse_response_text_data(response)
    
    def create_image(self, name:str, size:int, imageSetId:int, folderParentId:int=None, data:dict={}):
        """Creates a new image resource, and returns information for uploading the image to our servers for ingestion.
        Args:
            name (str): Label/name of the image.
            size (int): FileSize of the image.
            imageSetId (int): ID of the image set to create the image in.
            folderParentId (int, optional): ID of the folder to create the image in. Defaults to None.
            data (dict, optional): Other optional parameters. Defaults to {}.
                extractFoldersFromName (bool): Organizes images into folders based on the filename. If the appConfig forces this feature, then this parameter value will be ignored 
                cloneSourceId (str): ID of source images to clone for the new image. 
                clearMetadata (bool): Requires cloneSourceIds. Clears any cloned metadata from the duplicated image.
                storageSystemId (int): ID of storage system in which the image will be stored. If none specified, default is used.
                expandedBytes (int): summary size of the expanded image (populated when the original is a compressed file format)
                ingestedBytes (int): summary size of the data we write during ingestion and optimization.
        Returns:
            id (int): ID of the newly created image.
            objectPath (str): Path of the new image in S3.
        """
        data['name']=name
        data['size']=size
        data['imageSetId']=imageSetId
        if folderParentId:
            data['folderParentId']=folderParentId
        response = self.session.post(
            '{}/images'.format(self.endpoint),
            json=data
        )
        return self._handle_and_parse_response_text_data(response)

    def import_analysis_job_markups_from_file(self, analysis_job_id:int, annotations_xml):
        """Creates markups for an analysis job by parsing the contents of an XML file.
        old: import_markups_from_xml
        TODO different from documentation, report (endpoint, and required parameter)
        This API will put finishTime and Datereceive as if update_analysis_job with completed = True was called.
        Args:
            analysis_job_id (int): analysis job unique ID.
            annotations_xml (_type_): _description_
        """
        response = self.session.post(
            '{}/analysis/jobs/{}/markups/importFromXML'.format(self.endpoint, analysis_job_id),
            files={'files': StringIO(annotations_xml)}
        )
        return self._handle_and_parse_response_text_data(response)
    
    def import_image_annotations_from_file(self, imageId:int, annotations_xml):
        """ Creates annotations for an image by parsing the contents of an XML or MLD file.
        TODO Report: In documentation, this endpoint has no annotations_xml arguments.
        Args:
            imageId (int): image unique ID.
            annotations_xml ():
        """
        response = self.session.post(
            '{}/images/{}/annotations/import'.format(self.endpoint, imageId),
            files={'files': StringIO(annotations_xml)}
        )
        return self._handle_and_parse_response_text_data(response)
    
    ### Metadata
    def delete_metadata_field(self, fieldId:int):
        """Deletes properties of a metadata field.
        Args:
            fieldId (int): ID of the field.
        """
        response = self.session.delete('{}/metadata-fields/{}'.format(self.endpoint, fieldId))
        return self._handle_and_parse_response_text_data(response)
    
    def get_metadata_fields(self, filters:dict={}):
        """ Gets all metadata fields for a user.
        Without filters, it return all (no pagination).
        When putting 'imageSetId' as filter, other parms will be ignored (resourceType, fieldId). TODO report because it is not normal.
        Args:
            filters (dict, optional): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR).. Defaults to {}.
                imageSetId (list[int]): Filters by fields linked to an image set id. 
                organizationId (list[bool]): Filters by an organization id, or no organization id if set to null. 
                resourceType (list[str]): Filters results by whether they are a specific resource ('folder', 'image', 'imageSet', 'annotation', 'analysis').
                fieldId (list[int]): Filters results specific field ids.
                moduleId (list[int]): Filters results specific module ids.

        Returns:
            fields (dict):
                id (int): Unique id of the field.
                organizationId (int): Organization id of the field. if this exists, the field is a library field. 
                name (str): Name of the field.
                resourceType (str): Type of the field. Can be ('folder', 'image', 'imageSet', 'annotation', 'analysis'). 
                contentType (str): Type of the field's content. Can be 'string', 'hyperlink', 'number', 'dropdown', 'date', or 'boolean'. 
                addedBy (int): The id of the user who added the field. 
                addedByName (str): The name of the user who added the field. 
                requiredInOrganization (bool): Whether the field is required in its organization. 
                moduleId (int): The id of a module associated with the field. 
                created (date): The timestamp of the field's creation. 
                lastModified (date): The timestamp when the field was most recently modified.
        """
        response = self.session.get(
            '{}/metadata-fields'.format(self.endpoint),
            params={'filters': json.dumps(filters)}
        )
        return self._handle_and_parse_response_text_data(response)
    
    def get_metadata_values(self, filters:dict={}):
        """Gets all metadata values for which a user has access and match the filters specified.
        Gets all metadata fields for a user. Without filters, it return all (no pagination).
        Args:
            filters (dict, optional): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR). Defaults to {}.
                imageSetId (list[int]): Filters by fields linked to an image set id.
                imageId (list[int]): Filters fields by image fields linked to a image id.
                folderId (list[int]): Filters fields by folder fields linked to a folder id.
                annotationId (list[int]): Filters fields by annotation fields linked to a annotation id.
                analysisJobId (list[int]): Filters fields by analysis module fields linked to an analysis job id.
                organizationId (list[int]): Filters by an organization id, or no organization id if set to null.
                moduleId (list[int]): Filters by an module id, or no module id if set to null.
                resourceType (list[int]):Filters results by whether they are a specific resource ('folder', 'image', 'imageSet', 'annotation', 'analysis').
        Returns:
            metadataValues (list[dict]): An array of metadata values which match the filters. The following properties are of the object in the array, not the array itself.
                fieldId (int): The field for which the value is associated.
                resourceId (int): The specific uid for the given resource type that the value is associated.
                content (int/bool/string): The content of the value, whose type is determined by the field.
        """
        response = self.session.get(
            '{}/metadata-values'.format(self.endpoint),
            params={'filters': json.dumps(filters)}
        )
        return self._handle_and_parse_response_text_data(response)
    
    def update_metadata_field(self, field_id:int, data:dict={}):
        """Updates properties of a metadata field.
        Args:
            field_id (int): ID of the field.
            data (dict, optional): Defaults to {}.
                name (str): The name of the field.
                organizationId (int): The organization that uses this field.
                requiredInOrganization (bool): Whether the library field is required in its organization.
                dropdownOptions (list[dict]):The dropdown options for the field.
                    id (int): The id of the dropdown option 
                    moduleId (int): The id of a module associated with the field. 
                    value (str): The text of the dropdown option
                newDropdownOptions (list[str]): An array of new dropdown options to add to the dropdown field.
        """
        response = self.session.patch(
            '{}/metadata-fields/{}'.format(self.endpoint, field_id),
            json=data
        )
        return self._handle_and_parse_response_text_data(response)
    
    def update_metadata_values(self, metadata_values):
        """Updates the values of many metadata fields.
        Args:
            metadata_values (list[dict]): The values of metadata to be updated. The following properties are of the objects in the array, not the array itself.
                fieldId (int): The field for which the value is associated.
                resourceId (int): The specific uid for the given resource type that the value is associated.
                content (Any): The content for the field value.
        """
        response = self.session.patch(
            '{}/metadata-values'.format(self.endpoint),
            json={'metadataValues': metadata_values}
        )
        return self._handle_and_parse_response_text_data(response)

    def create_metadata_field(self, name:str, resourceType:str, contentType:str, organizationId:int, analysisModuleId:int=None, params:dict={}):
        """Creates properties of a metadata field.
        Args:
            name (str): The name of the field.
            resourceType (str): Type of the field. Can be ('folder', 'image', 'imageSet', 'annotation', 'analysis').
            contentType (str): Type of the field's content. Can be 'string', 'hyperlink', 'number', 'dropdown', 'date', or 'boolean'.
            organizationId (int): The organization that uses this field. Can be null for a custom field.
            analysisModuleId (int): A module to link to this field. This is not necessary to create a field, but without it, the field will not appear on concentriq if the resourceType is 'analysis'. TODO ask if normal.
            params (dict, optional): other optional parameters. Defaults to {}.
                imageSetId (int): A single initial image set to link this field into
                requiredInOrganization (bool): Whether the library field is required in its organization.
                dropdownOptions (list[str]): Initial options for a dropdown content type field.
                displayedInDashboard (bool): Sets the visibility in the dashboard case list (clinical).
                searchableInDashboard (bool): Sets ability to search in the dashboard case list (clinical).
                orderNumber (int): Sets an order number for sorting.
        """
        params['name']=name
        params['resourceType']=resourceType
        params['contentType']=contentType
        params['organizationId']=organizationId
        if analysisModuleId:
            params['analysisModuleId']=analysisModuleId
        response = self.session.post(
            '{}/metadata-fields'.format(self.endpoint),
            json=params)
        return self._handle_and_parse_response_text_data(response)

    ### Audit
    def get_repository_audit(self, imagesetId:int, dateStart=None, dateEnd=None):
        """ Gets the audit records for a single repository. Note: if one of dateStart or dateEnd is provided, both must be provided (if not all date will be fetch). A failure to provide valid dateStart/dateEnd parameters will return an error relating to the 'dateRange'.
        Args:
            imagesetId (int): ID of the imageSet to retrieve records for
            dateStart (str or datetime, optional): Start date to fetch records from.
            If str, "YYYY-MM-DD" format. Defaults to None.
            dateEnd (str or datetime, optional): End date to fetch records up to. 
            If str, "YYYY-MM-DD" format. This date will not be take in account. Defaults to None.
        """
        response = self.session.get(
            '{}/records/imageSets/{}'.format(self.endpoint, imagesetId),
            params={
                'dateStart': dateStart,
                'dateEnd': dateEnd
            }
        )
        return self._handle_response(response)

    def get_system_audit(self):
        """Gets the audit records for the system, excluding repository-level data.
        Returns:
            system-logs.csv (str): The requested audit records.
        """
        response = self.session.get('{}/records/system'.format(self.endpoint))
        return self._handle_response(response)

    ### Other APIs
    def get_annotations_for_image_xml(self, image_id:int):
        response = self.session.get('{}/images/{}/annotations/export/xml'.format(self.endpoint,image_id))
        return self._handle_and_parse_response_text(response)
    
    # ImageSet
    def get_imageset(self, imageSetId:int):
        """Returns the requested image set metadata.
        Args:
            imageSetId (int): image set unique ID.
        Returns:
            id (int): ID of the image set. 
            thumbnailURL (str): URL of the thumbnail for this image. 
            sharedWithPublic (bool): Describes whether the image set is public or not. 
            isFavorite (bool): Describes whether the image set has been favorited by the requesting user. 
            name (str): The name of the image set. 
            created (date): The date that the image set was created. 
            lastModified (date): The most recent date that the image set was modified. 
            imageCount (int): The Number of images currently in the image set. 
            totalSize (int): The total size in bytes of the image set. 
            ownerName (str): Name of the user who owns the image set. 
            ownerId (int): ID of the user that owns the image set. 
            description (str): The description of the image set. 
            groupId (int): ID of the group this image set belongs to, or null. 
            groupName (str): Name of the group this image set belongs to, or null. 
            sharePermissions (dict):  Permissions granted to the user who called this endpoint for the image set.
        """
        response = self.session.get('{}/imageSets/{}'.format(self.endpoint, imageSetId))
        return self._handle_and_parse_response_text_data(response)
    
    def get_imagesets(self, filters:dict={}, pagination:dict={}, includeMetadata:bool=None):
        """Returns all image sets for which the authenticated user has access.
        Args:
            filters (dict, optional): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR).Defaults to {}.
                ownerId (list[int]): An array of ownerIds to filter by the owner of the image set. 
                public (bool): Filters results by whether they are public or not. 
                isFavorite (bool): Filters results by whether they are favorited by this user or not. 
                isShared (bool): Filters results to only those which you have access to, but are not owned by you. 
                imageSetId (bool): Filters results by specific imageSetIds. 
                generalSearch (list[str]): An array of search phrases to search over various fields and properties. 
                name (list[str]): An array of names to filter by image set name. 
                description (list[str]): An array of descriptions to filter by image set descriptions. 
                fields (dict): An map with the metadataFieldId as the key and the value matching the contentType. Examples can be found above. 
                assignedUserId (list[int]): An array of userIds to filter by assigned user. 
                workflowId (list[int]): An array of workflowIds to filter by workflow. 
                stageId (list[int]): An array of stageIds to filter by stage. 
                organizationId (list[int]): An array of organizationIds to filter by 
                status (list[str]): An array of statuses to filter by stage status (case insensitive). 
                archived (bool): Determines whether to filter by imageSets that archived. 
                created (list[dict]): Filters results by the date the imageSet was created 
                    start (str): Start date for the query in ISO format 
                    end (str): End date for the query in ISO format 
                imageCount (list[dict]): An array of number field value objects 
                channelGroupName (str): Channel group name
            pagination (dict, optional): Options to specify sort order and size of subset for all image sets to return. Defaults to {}.
                rowsPerPage (int): How many rows to return. 
                page (int): Specifies an offset to advance the subset through the set of all image sets. 
                sortBy (list[str]): Which property of the case object to sort the entire set by. Options are ['name'], ['created'], ['lastModified'], ['size']. Also accepts a non-array String. To sort by a particular metadata value, use 'field_'. Ex: 'field_4'. 
                descending (bool): Specifies which order the values are sorted into
            includeMetadata (bool, optional): Optionally adds metadata values to the results. Defaults to None.

        Returns:
            image (list[dict)]): sets Array of image set objects. The following properties are of the objects in the array, not the array itself. 
                id (int): ID of the image set.
                thumbnailURL (str): URL of the thumbnail for this image.
                sharedWithPublic (bool): Describes whether the image set is public or not.
                isFavorite (bool): Describes whether the image set has been favorited by the requesting user.
                name (str): The name of the image set.
                created (date): The date that the image set was created.
                lastModified (date): The most recent date that the image set was modified.
                imageCount (int): The Number of images currently in the image set.
                totalSize (int): The total size in bytes of the image set.
                ownerName (str): Name of the user who owns the image set.
                ownerId (int): ID of the user that owns the image set.
                description (str): The description of the image set.
                groupId (int): ID of the group this image set belongs to, or null.
                groupName (str): Name of the group this image set belongs to, or null.
                sharePermissions (dict): Permissions granted to the user who called this endpoint for the image set.
                metadata (dict): Metadata keyed by the fieldId with the content as the value (if the 'includeMetadata' queryParameter is provided).
        """
        params={
            'filters' : json.dumps(filters),
            'pagination': json.dumps(pagination)
        }
        if includeMetadata:
            params['includeMetadata'] = includeMetadata
        response = self.session.get(
            '{}/imageSets'.format(self.endpoint),
            params=params
        )
        return self._handle_and_parse_response_text_data(response)
    
    def create_imageset(self, name:str, data:dict={}):
        """Creates a new image set and returns the newly created object.
        Args:
            name (str): Name of the new image set.
            data (dict, optional): Defaults to {}.
                groupId (int): ID of the group to assign the new image set to.
                templateId (int): ID of the template to provision fields with.
                stageId (int): ID of the initial stage this imageSet will have in a workflow.
                metadata (dict): ID of the initial stage this imageSet will have in a workflow.
        Returns:
            id (int): ID of the image set.
            thumbnailURL (str): URL of the thumbnail for this image.
            sharedWithPublic (bool): Describes whether the image set is public or not.
            isFavorite (bool): Describes whether the image set has been favorited by the requesting user.
            name (str): The name of the image set.
            created (date): The date that the image set was created.
            lastModified (date): The most recent date that the image set was modified.
            imageCount (int): The Number of images currently in the image set.
            totalSize (int): The total size in bytes of the image set.
            ownerName (str): Name of the user who owns the image set.
            ownerId (int): ID of the user that owns the image set.
            description (str): The description of the image set.
            groupId (int): ID of the group this image set belongs to, or null.
            groupName (str): Name of the group this image set belongs to, or null.
            sharePermissions (dict): Permissions granted to the user who called this endpoint for the image set.
        """
        data['name'] = name
        response = self.session.post(
            '{}/imageSets'.format(self.endpoint),
            json=data
        )
        return self._handle_and_parse_response_text_data(response)
    
    def delete_imageset(self, imageSetId:int):
        """Deletes the specified image set.
        Args:
            imageSetId (int): ID of the image set to be deleted.
        Returns:
            success (str): Message detailing the successful deletion of the image set.
        """
        response = self.session.delete('{}/imageSets/{}'.format(self.endpoint, imageSetId))
        return self._handle_and_parse_response_text_data(response)
    
    def export_imageset_csv(self, imageSetId:int):
        """Returns a file with all image set metadata as a CSV.
        Args:
            imageSetId (int): ID of the image set to be deleted.
        Returns:
            str: csv format
        """
        response = self.session.get('{}/imageSets/{}/export/csv'.format(self.endpoint, imageSetId))
        return self._handle_and_parse_response_text(response)
    
    # ImageSetGroups
    def get_imageset_groups(self, groupId:int):
        """Returns the group with the specified ID. The resulting object defines a Group.
        Args:
            groupId (int): ID of the group.
        Returns:
            id (str): Group Unique ID.
            name (str): Name of the image set group.
            imageSetCount (int): The number of image sets in the group.
            ownerName (str): Name of the owner of the image set group.
            ownerId (int): ID of the owner of the image set group.
            isFavorite (bool): Determines whether the image set group has been favorited by the requesting user.
            description (str): Description of the image set group.
            created (date): Date the image set group was created.
            lastModified (date): Date of the latest modification of the containing image sets.
            archived (bool): Determines whether the image set group is archived or not.
            sharePermissions (dict): Permissions granted to the user who called this endpoint for the group.
        """
        response = self.session.get(
            '{}/imageSetGroups/{}'.format(self.endpoint, groupId)
        )
        return self._handle_and_parse_response_text_data(response)
    
    def get_imagesets_groups(self, filters:dict={}, pagination:dict={}):
        """Returns the collection of image set groups to which you belong or own.
        Args:
            filters (dict, optional): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR). Defaults to {}.
                ownerId (list[int]): An array of specific ownerIds to filter by the owner of the resource.
                imageSetGroupId	(bool): Filters by specific Ids. 
                isFavorite (bool): Filters results by whether they are favorited by this user or not. 
                isShared (bool): Filters results by whether this resource has been shared with this user or not. Does not include resources owned by this user. 
                generalSearch (list[str]): An array of search phrases to search over various fields and case properties. 
                name (list[str]): An array of names to filter by image set group name. 
                created (list[dict]): Filters results by the date the imageSetGroup was created 
                    start (str): Start date for the query in ISO format 
                    end (str): End date for the query in ISO format
            pagination (dict, optional): Options to specify sort order and size of subset for all image sets to return. Defaults to {}.
                rowsPerPage (int): How many rows to return. 
                page (int): Specifies an offset to advance the subset through the set of all image sets. 
                sortBy (list[str]): Which property of the case object to sort the entire set by. Options are ['name'], ['created'], ['lastModified'], ['size']. Also accepts a non-array String. 
                descending (bool): pecifies which order the values are sorted into
        Returns:
            groups (list[dict]): Array of ImageSetGroups objects. The following properties are of the objects in the array, not the array itself. 
                id (str): Group Unique ID. 
                name (str): Name of the image set group. 
                imageSetCount (int): The number of image sets in the group. 
                ownerName (str): Name of the owner of the image set group. 
                ownerId (int): ID of the owner of the image set group. 
                description (str): Description of the image set group. 
                created (date): Date the image set group was created. 
                lastModified (date): Date of the latest modification of the containing image sets. 
                isFavorite (bool): Determines whether the image set group has been favorited by the requesting user. archived (bool): Determines whether the image set group is archived or not. 
                sharePermissions (object): Permissions granted to the user who called this endpoint for the group.
        """
        response = self.session.get(
            '{}/imageSetGroups'.format(self.endpoint),
            params={
                'filters' : json.dumps(filters),
                'pagination': json.dumps(pagination)
            }
        )
        return self._handle_and_parse_response_text_data(response)
    
    ### Templates
    # TODO report: return a dictionary with id as key and not a list of templates
    def get_templates(self, filters:dict={}):
        """Gets all Templates for a user.
        Args:
            filters (dict, optional): Object which specifies various filters. Each distinct filter can be considered as exclusive against one another (AND). Each item in an array can be considered inclusive against one another (OR). Defaults to {}.
                imageSetId (list[int]): Filters by templates linked to an image set id. 
                organizationId (list[int]): Filters by an organization id, or no organization id if set to null. 
                templateId (list[int]): Filters results by a specific template id 
                moduleId (list[int]): Filters results by a specific module id
        Returns:
            templates (list[dict]):
                id (int): Unique id of the template. 
                name (str): The name of the template. 
                description (str): The description of the template.
                ownerId (int): The owner of the template. 
                ownerName (str): The name of the template's owner. 
                moduleId (int): A module associated with the template.
                isDefault (bool): Whether the template is the default template for its owner's organization. 
                lastModified (date): The date that the template was last modified. 
                created (date): The date that the template was created.
        """
        response = self.session.get(
            '{}/templates'.format(self.endpoint),
            params={'filters' : json.dumps(filters)}
        )
        return self._handle_and_parse_response_text_data(response)
    
    ### ImageFluorescenceChannels
    # TODO report: biomarker and channel_index is not correctly documented
    def update_image_fluorescence_channel(self, image_id:int, channel_index:int, biomarker:str):
        """Updates image fluorescence channel properties. [imageId, channelIndex, biomarker, color, histogramMin, histogramMax].
        Args:
            image_id (int): image unique ID.
            channel_index (int): index of this fluorescent channel.
            biomarker (str): biomarker to be set for this channel.
        Returns:
            success (str): Message detailing the successful update of the image.
        """
        response = self.session.patch(
            '{}/imageFluorescenceChannels/{}/{}'.format(self.endpoint, image_id, channel_index),
            json={'biomarker': biomarker}
        )
        return self._handle_and_parse_response_text_data(response)
    
    ### SavedSearches
    # TODO report: documentation is savedSearch but it is savedSearches in endpoint
    def delete_saved_search(self, savedSearchId:int):
        """Deletes a saved search.
        Args:
            savedSearchId (int): ID of the savedSearch to delete.
        Returns:
            success (str): Message detailing the successful deletion of the saved search.
        """
        response = self.session.delete('{}/savedSearches/{}'.format(self.endpoint, savedSearchId))
        return self._handle_and_parse_response_text_data(response)
    
    def get_saved_searches(self):
        """Returns all saved searches accessible to the user.
        Returns:
            savedSearches (list[dict]): Containing object for the savedSearch objects. The following properties are of the objects in the array, not the array itself. 
                id (int): ID of the savedSearch. 
                query (dict): key value mappings to evaluate as a search query. 
                name (str): name of the savedSearch. 
                createdBy (str): the user who created the savedSearch. 
                createdAt (str): the time when the savedSearch was created. 
                lastUpdatedAt (str): the time when the savedSearch was last updated.
        """
        response = self.session.get('{}/savedSearches'.format(self.endpoint))
        return self._handle_and_parse_response_text_data(response)
    
    def update_saved_search(self, savedSearchId:int, name:str):
        """Updates a saved search.
        Args:
            savedSearchId (int): ID of the savedSearch to update.
            name (str): new name for the savedSearch.
        Returns:
            savedSearch (dict): the savedSearch object 
                id (int): ID of the savedSearch. 
                query (dict): key value mappings to evaluate as a search query. 
                name (str): name of the savedSearch.
                createdBy (str): the user who created the savedSearch. 
                createdAt (str): the time when the savedSearch was created.
                lastUpdatedAt (str): the time when the savedSearch was last updated.
        """
        response = self.session.patch(
            '{}/savedSearches/{}'.format(self.endpoint, savedSearchId),
            json={'name': name}
        )
        return self._handle_and_parse_response_text_data(response)
    
    def create_saved_search(self, query:dict, name:str):
        """Saves a search for the user.
        Args:
            query (Object): key value mappings to evaluate as a search query.
            name (str): name of the savedSearches.
        Returns:
            savedSearch (dict): The savedSearch object 
                id (int): ID of the savedSearch. 
                query (dict): key value mappings to evaluate as a search query. 
                name (str): name of the savedSearch. 
                createdBy (str): the user who created the savedSearch.
                createdAt (str): the time when the savedSearch was created. 
                lastUpdatedAt (str): the time when the savedSearch was last updated.
        """
        response = self.session.post(
            '{}/savedSearches'.format(self.endpoint),
            json={
                'query':query,
                'name':name
            }
        )
        return self._handle_and_parse_response_text_data(response)
    
    ### AnnotationClasses
    def get_annotation_classes(self, filters:dict={}):
        """Fetches all annotationClasses.
        Args:
            filters (dict, optional): Builds a paginated result set. Allows you to filter, order and paginate your results. Defaults to {}.
                fields (list[str]): Fields from the base resourceType and related resources that should be returned in the query. If not defined, all default fields are returned. 
                limit (int): >= 0, The maximum number of items to return in a paginated set. 
                offset (int): >= 0, The number of items to shift the paginated results by. 
                order (list[dict]): An ordered array of properties and the direction to sort them by. 
                eager (dict): QueryFilterEagerExpression
            total (bool): Include total results count in response
        Returns:
            items (list[dict]): AnnotationClass 
                id (int): The UID of the resource 
                name (str): >= 2 characters, Annotation class name 
                description (str): Annotation class description 
                color (str^#[0-9A-Fa-f]{6}$) Annotation class color represented as a hexadecimal string 
                createdAt (str): <date-time>
                createdBy (int)
                lastUpdatedAt (str): <date-time> 
                lastUpdatedBy (int)
                imageSets (list[dict]): ImageSet
                annotations (list[dict]): Annotation
                creator (dict): User
            total (int)
        """
        response = self.session.get(
            '{}/v3/annotationClasses'.format(self.endpoint),
            params={
                'filter': json.dumps(filters)
            }
        )
        return self._handle_and_parse_response_text_json(response)
        
    def create_annotation_class(self, name:str, color:str, description:str=None):
        """Creates a annotationClass.
        Args:
            name (str): Annotation class name.
            color (str): Annotation class color represented as a hexadecimal string.
            description (str, optional): Annotation class description. Defaults to None.

        Returns:
            id (int): The UID of the resource 
            name (str): >= 2, characters Annotation class name 
            description (str): Annotation class description
            color (str^#[0-9A-Fa-f]{6}$) Annotation class color represented as a hexadecimal string 
            createdAt (str): <date-time> 
            createdBy (int) 
            lastUpdatedAt (str): <date-time> 
            lastUpdatedBy (int)
        """
        data = {
            'name': name,
            'color': color,
        }
        if description:
            data['description'] = description
        response = self.session.post(
            '{}/v3/annotationClasses'.format(self.endpoint),
            json=data
        )
        return self._handle_and_parse_response_text_json(response)
        
    def update_annotation_class(self, annotationClassId:int, name:str=None, color:str=None, description:str=None):
        """
        Args:
            annotationClassId (int): An ID of a annotationClass.
            name (str, optional): Annotation class name. Defaults to None.
            color (str, optional): Annotation class color represented as a hexadecimal string. Defaults to None.
            description (str, optional): Annotation class description. Defaults to None.
        Returns:
            id (int): The UID of the resource 
            name (str): >= 2 characters Annotation class name 
            description (str): Annotation class description 
            color (str^#[0-9A-Fa-f]{6}$): Annotation class color represented as a hexadecimal string 
            createdAt (str): <date-time> 
            createdBy (int) 
            lastUpdatedAt (str): <date-time> 
            lastUpdatedBy (int)
        """
        data = {}
        if name:
            data['name'] = name
        if color:
            data['color'] = color
        if description:
            data['description'] = description
        response = self.session.patch(
            '{}/v3/annotationClasses/{}'.format(self.endpoint, annotationClassId),
            json=data
        )
        return self._handle_and_parse_response_text_json(response)
        
    def delete_annotation_class(self, annotationClassId:int):
        """
        Args:
            annotationClassId (int): An ID of a annotationClass.
        Returns:
            str: The annotationClass was successfully deleted
        """
        response = self.session.delete('{}/v3/annotationClasses/{}'.format(self.endpoint, annotationClassId))
        return self._handle_and_parse_response_text(response)
        
    def get_annotation_class(self, annotationClassId:int):
        """
        Args:
            annotationClassId (int): An ID of a annotationClass.
        Returns:
            id (int): The UID of the resource
            name (str): >= 2 characters Annotation class name
            description (str):  Annotation class description
            color (str^#[0-9A-Fa-f]{6}$): Annotation class color represented as a hexadecimal string
            createdAt (str): <date-time>
            createdBy (int)
            lastUpdatedAt (str): <date-time>
            lastUpdatedBy (int)
        """
        response = self.session.get('{}/v3/annotationClasses/{}'.format(self.endpoint, annotationClassId))
        return self._handle_and_parse_response_text_json(response)
    
    def add_annotation_class_to_imageset(self, annotationClassId:int, imageSetId:int):
        """Add an annotation class to an image set.
        Args:
            annotationClassId (int): An ID of a annotationClass.
            imageSetId (int): An ID of a imageSet.
        Returns:
            success (str): "Annotation class has been added to the image set"
        """
        response = self.session.post('{}/v3/annotationClasses/{}/imageSets/{}'.format(self.endpoint, annotationClassId, imageSetId))
        return self._handle_and_parse_response_text(response)
        
    def remove_annotation_class_from_imageset(self, annotationClassId:int, imageSetId:int):
        """Remove an annotation class from an image set it has been added to.
        Args:
            annotationClassId (int): An ID of a annotationClass.
            imageSetId (int): An ID of a imageSet.
        Returns:
            success (str): "Annotation class has been removed from the image set"
        """
        response = self.session.delete('{}/v3/annotationClasses/{}/imageSets/{}'.format(self.endpoint, annotationClassId, imageSetId))
        return self._handle_and_parse_response_text(response)