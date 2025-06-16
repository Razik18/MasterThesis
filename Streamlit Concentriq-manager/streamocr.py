import streamlit as st
import requests
import base64
import json
import os
import win32cred
import pandas as pd
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder
from paddleocr import PaddleOCR
import pandas as pd



API_URL = ""
PROXY = {
    "http": "",
    "https": "",
}

settings_file = os.path.join(os.path.dirname(__file__), "settings.json")

try:
    with open(settings_file, "r") as file:
        settings = json.load(file)  
        FIELD_IDS = settings.get("FIELD_IDS", [])  
        accession_row = settings.get("accession_row", 1)  
        print(f"Loaded FIELD_IDS: {FIELD_IDS}")
        print(f"Loaded accession_row: {accession_row}")
except FileNotFoundError:
    print(f"Error: {settings_file} not found. Please ensure the file exists in the same directory as this script.")
except json.JSONDecodeError as e:
    print(f"Error decoding JSON in {settings_file}: {e}")


def get_credentials():
    """Retrieve credentials from Windows Credential Manager."""
    try:
        cred = win32cred.CredRead("ConcentriqProd", win32cred.CRED_TYPE_GENERIC)
        username = cred['UserName']
        password = cred['CredentialBlob'].decode('utf-8').replace('\x00', '').strip()
        return username, password
    except Exception as e:
        print(f"Error retrieving credentials: {e}")
        return None, None

def authenticate():
    """Generate Basic Authentication headers."""
    EMAIL, PASSWORD = get_credentials()
    credentials = f"{EMAIL}:{PASSWORD}"
    encoded_credentials = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")
    return {"Authorization": f"Basic {encoded_credentials}"}

def get_image_url(image_id):
    """Retrieve the signed URL for the label image using the image ID."""
    headers = authenticate()
    if not headers:
        return None
    try:
        response = requests.get(f"{API_URL}/images/{image_id}", headers=headers, proxies=PROXY)
        if response.status_code == 200:
            data = response.json()
            associated_images = data.get("data", {}).get("associatedImages", [])
            for image in associated_images:
                if image["type"] == "label":
                    return image["signedURL"]
            st.error("Label image not found.")
        else:
            st.error(f"Failed to fetch image: {response.status_code} - {response.text}")
    except Exception as e:
        st.error(f"Error retrieving image URL: {e}")
    return None

def get_macro_url(image_id):
    """Retrieve the signed URL for the label image using the image ID."""
    headers = authenticate()
    if not headers:
        return None
    try:
        response = requests.get(f"{API_URL}/images/{image_id}", headers=headers, proxies=PROXY)
        if response.status_code == 200:
            data = response.json()
            associated_images = data.get("data", {}).get("associatedImages", [])
            for image in associated_images:
                if image["type"] == "macro":
                    return image["signedURL"]
            st.error("Macro image not found.")
        else:
            st.error(f"Failed to fetch image: {response.status_code} - {response.text}")
    except Exception as e:
        st.error(f"Error retrieving image URL: {e}")
    return None

def perform_ocr(image_url):
    """Download the image and perform OCR using PaddleOCR."""
    try:
        response = requests.get(image_url, proxies=PROXY, stream=True)
        if response.status_code == 200:
            ocr = PaddleOCR(use_angle_cls=True, lang='en')
            result = ocr.ocr(response.content, cls=True)
            print(result)
            extracted_text = "\n".join([line[1][0] for line in result[0]])
            return extracted_text
        else:
            st.error(f"Failed to download image: {response.status_code} - {response.text}")
    except Exception as e:
        st.error(f"Error performing OCR: {e}")
    return None

class ConcentriqAPI:
    def __init__(self):
        self.endpoint = API_URL
        self.session = requests.Session()
        self.session.headers.update(authenticate())

    def update_metadata_values(self, metadata_values):
        """Updates the values of many metadata fields.
        Args:
            metadata_values (list[dict]): The values of metadata to be updated. The following properties are of the objects in the array, not the array itself.
                fieldId (int): The field for which the value is associated.
                resourceId (int): The specific uid for the given resource type that the value is associated.
                content (Any): The content for the field value.
        """
        response = self.session.patch(
            f'{self.endpoint}/metadata-values',
            json={'metadataValues': metadata_values}
        )
        if response.status_code == 200:
            st.success("Metadata updated successfully.")
        else:
            st.error(f"Failed to update metadata. Response: {response.text}")
    def get_metadata_field_options(self, field_id):
        """Fetch the dropdown options for a specific metadata field using filters."""
        filters = {"fieldId": [field_id]}  
        
        try:
            response = self.session.get(
                f'{self.endpoint}/metadata-fields',
                params={'filters': json.dumps(filters)}
            )
            if response.status_code == 200:
                return response.json()
            else:
                raise Exception(f"Failed to fetch metadata field options: {response.text}")
        except requests.exceptions.RequestException as e:
            raise Exception(f"An error occurred while fetching metadata field options: {e}")

    def update_image_name(self, image_id, new_name):
        """Update the image name using a PATCH request."""
        update_url = f"{self.endpoint}/images/{image_id}"
        payload = {"name": new_name}

        try:
            response = self.session.patch(update_url, json=payload, proxies=PROXY)
            if response.status_code == 200:
                st.success(f"Image ID {image_id} renamed to {new_name}.")
            else:
                st.error(f"Failed to rename Image ID {image_id}. Response: {response.text}")
        except requests.exceptions.RequestException as e:
            st.error(f"An error occurred while renaming the image: {e}")

# Fetch metadata from API
def get_metadata():
    url = f"{API_URL}metadata-fields"
    try:
        response = requests.get(url, headers=authenticate())
        response.raise_for_status()
        metadata = response.json()
        fields = metadata.get("data", {}).get("fields", [])
        return [{"name": field["name"], "id": field["id"]} for field in fields]
    except Exception as e:
        st.error(f"Failed to fetch metadata: {e}")
        return []


def load_field_ids():
    try:
        with open(settings_file, "r") as file:
            settings = json.load(file)
            return settings.get("FIELD_IDS", [])
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def save_field_ids(field_ids):
    try:
        with open(settings_file, "w") as file:
            json.dump({"FIELD_IDS": field_ids}, file, indent=4)
        return True
    except Exception as e:
        st.error(f"Error saving FIELD_IDS: {e}")
        return False


def get_all_images_by_imageset_id(imageset_id):
    """Retrieve all images associated with a specific Repository ID."""
    headers = authenticate()
    images_url = f"{API_URL}images"
    filters = {"imageSetId": [imageset_id]}
    
    page = 1
    all_images = []

    while True:
        params = {
            "filters": json.dumps(filters),
            "pagination": json.dumps({"rowsPerPage": 100, "page": page}),
        }

        try:
            response = requests.get(images_url, headers=headers, params=params, proxies=PROXY)
            if response.status_code == 200:
                data = response.json().get("data", {})
                images = data.get("images", [])
                if not images:
                    break  
                all_images.extend(images)
                page += 1
            else:
                st.error(f"Failed to retrieve images for Repository ID {imageset_id}. Response: {response.text}")
                break
        except requests.exceptions.RequestException as e:
            st.error(f"An error occurred while retrieving images: {e}")
            break

    return all_images

def get_metadata_for_images(image_ids):
    """Retrieve metadata for multiple images in batches."""
    headers = authenticate()
    metadata_url = f"{API_URL}metadata-values"
    batch_size = 50  
    all_metadata = []

    for i in range(0, len(image_ids), batch_size):
        batch_ids = image_ids[i:i+batch_size]
        filters = {"imageId": batch_ids}
        params = {"filters": json.dumps(filters)}

        try:
            response = requests.get(metadata_url, headers=headers, params=params, proxies=PROXY)
            if response.status_code == 200:
                metadata = response.json().get("data", [])
                all_metadata.extend(metadata)
            else:
                st.error(f"Failed to retrieve metadata for batch {batch_ids}. Response: {response.text}")
        except requests.exceptions.RequestException as e:
            st.error(f"An error occurred while retrieving metadata for batch {batch_ids}: {e}")

    return all_metadata


def get_folders_by_imageset_id(imageset_id, block_id):
    """Retrieve folders associated with a specific Repository ID, filtered by name."""
    print(f"Filtering by imageSetId: {imageset_id} and name: {block_id}")
    
    headers = authenticate()
    folders_url = f"{API_URL}folders"
    
    filters = {
        "imageSetId": [imageset_id],  
        "name": [block_id]            
    }
    
    params = {
        "filters": json.dumps(filters),  
        "idsOnly": False  
    }

    try:
        response = requests.get(folders_url, headers=headers, params=params, proxies=PROXY)
        
        if response.status_code == 200:
            folders_data = response.json().get("data", [])
            return folders_data  
        else:
            st.error(f"Failed to retrieve folders for Repository ID {imageset_id}. Response: {response.text}")
            return []
    except requests.exceptions.RequestException as e:
        st.error(f"An error occurred while retrieving folders: {e}")
        return []



def update_image_name(image_id, new_name):
    """Update the image name using a PATCH request."""
    headers = authenticate()
    update_url = f"{API_URL}images/{image_id}"
    payload = {"name": new_name}

    try:
        response = requests.patch(update_url, json=payload, headers=headers, proxies=PROXY)
        if response.status_code == 200:
            st.success(f"Image ID {image_id} renamed to {new_name}.")
            return True
        else:
            st.error(f"Failed to rename Image ID {image_id}. Response: {response.text}")
            return False
    except requests.exceptions.RequestException as e:
        st.error(f"An error occurred while renaming the image: {e}")
        return False



def get_field_names_by_ids(field_ids):
    """Retrieve field names based on field IDs."""
    headers = authenticate()  
    metadata_url = f"{API_URL}metadata-fields"
    
    try:
        response = requests.get(metadata_url, headers=headers, proxies=PROXY)
        if response.status_code == 200:
            data = response.json().get("data", {})
            fields = data.get("fields", [])

            if not isinstance(fields, list):
                st.error("Unexpected response format: 'fields' is not a list.")
                return {}

            field_names = {field["id"]: field["name"] for field in fields if field["id"] in field_ids}

            ordered_field_names = {field_id: field_names.get(field_id) for field_id in field_ids if field_id in field_names}
            
            return ordered_field_names
        else:
            st.error(f"Failed to retrieve metadata fields. Response: {response.text}")
            return {}
    except requests.exceptions.RequestException as e:
        st.error(f"An error occurred while retrieving metadata fields: {e}")
        return {}

 
def generate_new_image_name_from_fields(selected_rows, image_id, ordered_field_names):
    new_name_parts = []
    
    selected_row = selected_rows[selected_rows['Image ID'] == image_id].iloc[0]
    
    for field_name in ordered_field_names:
        if field_name in selected_row:
            content = selected_row[field_name]
            if content:  
                new_name_parts.append(content)
            else:
                st.warning(f"Warning: Content for {field_name} is empty.")
        else:
            st.warning(f"Field {field_name} not found in selected_row.")
    

    if not new_name_parts:
        st.error(f"Error: No valid fields found to construct image name for Image ID {image_id}.")
        return "Unnamed_Image"

    return "_".join(new_name_parts)






st.set_page_config(
    page_title="OCR",  
    layout="wide",  
    initial_sidebar_state="expanded"  
)

selected_page = st.sidebar.selectbox("Navigation", ["Home", "Settings"])
if selected_page == "Home":
    st.title("Manage Images in Repository")
    st.write("Select images and extract labels with OCR")
    api_client = ConcentriqAPI()
    field_names = get_field_names_by_ids(FIELD_IDS)
    print(field_names)
    print(FIELD_IDS)


    if "image_data" not in st.session_state:
        st.session_state.image_data = None
    if "df" not in st.session_state:
        st.session_state.df = None
    if "previous_repository_id" not in st.session_state:
        st.session_state.previous_repository_id = ""


    if field_names:
        rename_pattern = "_".join(field_names.values())
    else:
        st.write("Field names could not be retrieved.")

    col1, col2 = st.columns([2, 1])
    with col1:
        source_repository_id = st.text_input("Enter Source Repository ID:")
    if source_repository_id != st.session_state.previous_repository_id:
        st.session_state.image_data = None
        st.session_state.df = None
        st.session_state.previous_repository_id = source_repository_id






    load_images = st.checkbox("Display Repository Images")
    if load_images:
        if st.session_state.image_data is None:
            images = get_all_images_by_imageset_id(source_repository_id)
            if images:
                image_ids = [image.get("id") for image in images]
                metadata_list = get_metadata_for_images(image_ids)  
                metadata_mapping = {}
                for item in metadata_list:
                    resource_id = item.get("resourceId")  
                    if resource_id not in metadata_mapping:
                        metadata_mapping[resource_id] = []
                    metadata_mapping[resource_id].append(item)

                image_data = []
                for image in images:
                    image_id = image.get("id")
                    created_date = image.get("created")
                    date = datetime.strptime(created_date, "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d-%m-%Y") if created_date else None

                    field_values = {"Image Name": image.get("name"), "Date": date, "Image ID": image_id}

                    metadata_for_image = metadata_mapping.get(image_id, [])  
                    for item in metadata_for_image:
                        field_id = item.get("fieldId")
                        if field_id in FIELD_IDS:  
                            field_name = field_names.get(field_id)
                            if field_name:  
                                field_values[field_name] = item.get("content")

                    image_data.append(field_values)

                st.session_state.image_data = image_data
                st.session_state.df = pd.DataFrame(image_data)

        if st.session_state.df is not None:
            gb = GridOptionsBuilder.from_dataframe(st.session_state.df)
            gb.configure_selection('multiple', use_checkbox=True)
            gb.configure_column('Image Name', headerCheckboxSelection = True)
            grid_options = gb.build()

            grid_response = AgGrid(st.session_state.df, gridOptions=grid_options, enable_enterprise_modules=False)

            selected_rows = grid_response["selected_rows"]

            load_ocr = st.checkbox("OCR read", key="load_ocr_checkbox")

            if load_ocr:
                if selected_rows is not None and len(selected_rows) > 0:
                    selected_indices = selected_rows['Image ID'].tolist()
                    st.session_state.selected_indices = selected_indices


                    for image_id in selected_indices:
                        if image_id: 
                            if f"ocr_processed_{image_id}" not in st.session_state:
                                image_url = get_image_url(image_id)
                                macro_url = get_macro_url(image_id)

                                if image_url:
                                    text = perform_ocr(image_url)

                                    st.session_state[f"ocr_processed_{image_id}"] = True
                                    st.session_state[f"ocr_text_{image_id}"] = text

                                else:
                                    st.error(f"No label found for Image ID {image_id}. Unable to fetch image.")
                            else:
                                text = st.session_state.get(f"ocr_text_{image_id}")
                                image_url = get_image_url(image_id)
                                macro_url = get_macro_url(image_id)


                            if text:
                                st.subheader(f"OCR Results for Image {image_id}")
                                col1, col2, col3, col4 = st.columns([3, 3, 3, 4])

                                with col1:
                                    st.image(macro_url, caption="Macro Image", use_container_width=True)
                                with col2:
                                    st.image(image_url, caption="Label Image", use_container_width=True)

                                with col3:
                                    st.text_area(f"Extracted Text for Image {image_id}", text, key=f"text_area_{image_id}", height=200)

                                with col4:
                                    stain_options = api_client.get_metadata_field_options(31)  
                                    if "data" in stain_options:
                                        stain_values = {
                                            option["id"]: option["value"] for option in stain_options["data"]["fields"][0].get("dropdownOptions", [])
                                        }
                                    else:
                                        stain_values = {}  

                                    stain_values_list = [""] + list(stain_values.values())

                                    row_index = accession_row - 1
                                    ocr_lines = text.split('\n')
                                    first_row_text = ocr_lines[row_index] if len(ocr_lines) > row_index else ""

                                    if first_row_text.startswith("QC") and not first_row_text.startswith("QC "):
                                        first_row_text = first_row_text.replace("QC", "QC ", 1)

                                    if f"accession_{image_id}" not in st.session_state:
                                        st.session_state[f"accession_{image_id}"] = first_row_text
                                    if f"stain_{image_id}" not in st.session_state:
                                        st.session_state[f"stain_{image_id}"] = ""
                                    if f"prefilled_name_{image_id}" not in st.session_state:
                                        st.session_state[f"prefilled_name_{image_id}"] = ""

                                    accession = st.text_input(
                                        f"Accession for Image {image_id}",
                                        value=st.session_state[f"accession_{image_id}"],
                                        key=f"accession_{image_id}"
                                    )
                                    stain = st.selectbox(
                                        f"Stain for Image {image_id}",
                                        options=stain_values_list,
                                        key=f"stain_{image_id}"
                                    )


                                    if accession and stain:
                                        st.session_state[f"prefilled_name_{image_id}"] = f"{accession} {stain}"
                                    else:
                                        st.session_state[f"prefilled_name_{image_id}"] = ""


                                    with st.form(key=f"big_form_{image_id}"):
                                        imagename = st.text_input(
                                            f"Image name for Image {image_id}",
                                            value=st.session_state[f"prefilled_name_{image_id}"], 
                                            key=f"user_provided_name_{image_id}"  
                                        )

                                        if st.form_submit_button(f"Submit Updates for Image {image_id}"):
                                            accession = st.session_state[f"accession_{image_id}"]
                                            stain = st.session_state[f"stain_{image_id}"]

                                            if accession:  
                                                api_client.update_metadata_values([{
                                                    "fieldId": 29,  
                                                    "resourceId": int(image_id),
                                                    "content": accession
                                                }])

                                            if stain: 
                                                selected_stain_id = next(
                                                    (key for key, value in stain_values.items() if value == stain),
                                                    None
                                                )
                                                if selected_stain_id:
                                                    api_client.update_metadata_values([{
                                                        "fieldId": 31,  
                                                        "resourceId": int(image_id),
                                                        "content": stain
                                                    }])

                                            user_image_name = st.session_state[f"user_provided_name_{image_id}"]  
                                            if user_image_name:  
                                                api_client.update_image_name(image_id, user_image_name)
                                            else:
                                                st.warning(f"No image name provided for Image {image_id}.")

                                            st.success(f"Updates for Image {image_id} have been applied.")


                                if not text:
                                    st.error(f"OCR text could not be extracted for Image ID {image_id}.")
                        else:
                            st.error(f"Failed to fetch the image for Image ID {image_id}.")
                else:
                    st.error("No images selected. Please select at least one image to process.")
            else:
                st.info("Please select 'OCR read' to process images.")
        else:
            st.error("No images found in the source repository.")
    else:
        st.info("Check the box above to load images.")
elif selected_page == "Settings":
    st.title("Settings")
    st.write("Update the FIELD_IDS and Accession Row here.")

    metadata = get_metadata()

    if metadata:
        dropdown_options = [f"{field['name']} (ID: {field['id']})" for field in metadata]


        selected_fields = st.multiselect(
            "Select Fields",
            options=dropdown_options,
            default=[],  
            help="Choose the fields you want to include in FIELD_IDS."
        )


        selected_ids = [
            field["id"]
            for field in metadata
            if f"{field['name']} (ID: {field['id']})" in selected_fields
        ]

        st.write("Selected FIELD_IDS:", selected_ids)

        accession_row_input = st.number_input(
            "Enter Accession Row (e.g., 1 for row 1):",
            min_value=1,  
            step=1,  
            value=settings.get("accession_row", 1)  
        )


        if st.button("Save Changes"):
            if save_field_ids(selected_ids):
                settings["accession_row"] = accession_row_input  
                with open(settings_file, "w") as file:
                    json.dump(settings, file, indent=4) 

                st.success("FIELD_IDS and Accession Row updated successfully!")
            else:
                st.error("Failed to update FIELD_IDS.")
    else:
        st.error("No metadata available. Please check the API connection.")


