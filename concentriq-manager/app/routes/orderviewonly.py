from app.config import GLOBAL_STATE

from flask import (
    Blueprint, render_template, redirect, url_for, render_template, request, flash
)

orderviewonly = Blueprint('orderviewonly', __name__)

@orderviewonly.route("/orderviewonly")
def index():
    # If not logged, return to login page
    concentriq_service = GLOBAL_STATE['service']
    if concentriq_service is None:
        return redirect(url_for("login.index"))
    
    metadata_fields = None
    
    return render_template(
        'orderviewonly_page.html',
        metadata_fields=metadata_fields
    )

@orderviewonly.route("/orderviewonly/updatetable", methods=["POST"])
def updatetable():
    # If not logged, return to login page
    concentriq_service = GLOBAL_STATE['service']
    if concentriq_service is None:
        return redirect(url_for("login.index"))
    
    repository_id = request.form.get("source_repository_id")
    valid_repo, repo_info = check_valid_repo(concentriq_service, repository_id)
    if not valid_repo:
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    # Get metadata fields
    metadata_fields = get_metadata_fields(concentriq_service, repository_id)

    # Filter metadata fields depending on pushed button, and fill table
    if request.method == 'POST':
        if request.form.get('loadImageField') == 'Load Image fields':
            filtered_metadata_fields = [metadata_field for metadata_field in metadata_fields if metadata_field.get('content_type') == 'image']
        elif  request.form.get('loadFolderField') == 'Load Folder Fields':
            filtered_metadata_fields = [metadata_field for metadata_field in metadata_fields if metadata_field.get('content_type') == 'folder']
        elif  request.form.get('loadImageSetField') == 'Load ImageSet Fields':
            filtered_metadata_fields = [metadata_field for metadata_field in metadata_fields if metadata_field.get('content_type') == 'imageSet']
        else:
            flash("error to load metadata fields.", "error")
            return render_template("orderviewonly.index")
    
    filtered_metadata_fields = sorted(filtered_metadata_fields, key=lambda x: x['order'])
    return render_template(
        'orderviewonly_page.html',
        repository_text=get_repository_text(repo_info),
        repository_id=repository_id,
        metadata_fields=filtered_metadata_fields
    )

@orderviewonly.route("/orderviewonly/esmOrderUpdate", methods=["POST"])
def esmOrderUpdate():
    # If not logged, return to login page
    concentriq_service = GLOBAL_STATE['service']
    if concentriq_service is None:
        return redirect(url_for("login.index"))
    
    repository_id = request.form.get("hiddenrepoid1")
    valid_repo, repo_info = check_valid_repo(concentriq_service, repository_id)
    if not valid_repo:
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    ## Update order of "eSM" containing fields to the end
    filters = {
        'imageSetId': [repository_id]
    }
    res = concentriq_service.client.get_metadata_fields(filters)
    end_order = len(res.get('fields')) + 1
    
    for field_ in res.get('fields'):
        if 'eSM' in field_.get('name'):
            res = update_imageset_properties(concentriq_service.client, repository_id, field_.get('id'), orderNumber=end_order)
            if res is None:
                print(f"!!! Error for field (id: {field_.get('id')} for imageset (id: {repository_id}) !!!")
                flash(f"!!! Error for field (id: {field_.get('id')} for imageset (id: {repository_id}) !!!", "error")
            else:
                print(f"Change to order ({end_order}) for: field (id: {field_.get('id')}) --- name: {field_.get('name')}")
    
    # Get metadata fields and filter by content_type
    metadata_fields = get_metadata_fields(concentriq_service, repository_id)
    filtered_metadata_fields = sorted(metadata_fields, key=lambda x: x['content_type'])
    
    return render_template(
        'orderviewonly_page.html',
        repository_text=get_repository_text(repo_info),
        repository_id=repository_id,
        metadata_fields=filtered_metadata_fields
    )

@orderviewonly.route("/orderviewonly/esmViewOnlyUpdate", methods=["POST"])
def esmViewOnlyUpdate():
    # If not logged, return to login page
    concentriq_service = GLOBAL_STATE['service']
    if concentriq_service is None:
        return redirect(url_for("login.index"))
    
    repository_id = request.form.get("hiddenrepoid2")
    valid_repo, repo_info = check_valid_repo(concentriq_service, repository_id)
    if not valid_repo:
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    # Update order of "eSM" containing fields to ViewOnly
    filters = {
        'imageSetId': [repository_id]
    }
    res = concentriq_service.client.get_metadata_fields(filters)
    
    for field_ in res.get('fields'):
        if 'eSM' in field_.get('name'):
            res = update_imageset_properties(concentriq_service.client, repository_id, field_.get('id'), viewOnly=True, study=False)
            if res is None:
                print(f"Error for field (id: {field_.get('id')} for imageset (id: {repository_id})")
                flash(f"Error for field (id: {field_.get('id')} for imageset (id: {repository_id})", "error")
            else:
                print(f"Change to view only for: field (id: {field_.get('id')}) --- name: {field_.get('name')}")

    # Get metadata fields and filter by content_type
    metadata_fields = get_metadata_fields(concentriq_service, repository_id)
    filtered_metadata_fields = sorted(metadata_fields, key=lambda x: x['content_type'])
    
    return render_template(
        'orderviewonly_page.html',
        repository_text=get_repository_text(repo_info),
        repository_id=repository_id,
        metadata_fields=filtered_metadata_fields
    )

@orderviewonly.route("/orderviewonly/UpdateViewOnly", methods=["POST"])
def UpdateViewOnly():
    # If not logged, return to login page
    concentriq_service = GLOBAL_STATE['service']
    if concentriq_service is None:
        return redirect(url_for("login.index"))
    
    data = request.json
    repository_id = data.get('repository_id')
    metadata_field_id = data.get('metadata_field_id')
    checkbox_checked = data.get('checkbox_checked')
    
    valid_repo, repo_info = check_valid_repo(concentriq_service, repository_id)
    if not valid_repo:
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    try:
        metadata_field_id = int(metadata_field_id)
    except Exception as e:
        flash(f"{metadata_field_id} Not convertable to integer. ({e})", "error")
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    # Update ViewOnly status according to checkbox
    res = update_imageset_properties(concentriq_service.client, repository_id, metadata_field_id, viewOnly=checkbox_checked, study=False)
    
    if res is None:
        print(f"Error for field (id: {metadata_field_id} for imageset (id: {repository_id})")
        flash(f"Error for field (id: {metadata_field_id} for imageset (id: {repository_id})", "error")
    else:
        print(f"Change to view only for: field (id: {metadata_field_id})")

    # Get metadata fields and filter by content_type
    metadata_fields = get_metadata_fields(concentriq_service, repository_id)
    
    return render_template(
        'orderviewonly_page.html',
        repository_text=get_repository_text(repo_info),
        repository_id=repository_id,
        metadata_fields=metadata_fields
    )

@orderviewonly.route("/orderviewonly/UpdateStudy", methods=["POST"])
def UpdateStudy():
    # If not logged, return to login page
    concentriq_service = GLOBAL_STATE['service']
    if concentriq_service is None:
        return redirect(url_for("login.index"))
    
    data = request.json
    repository_id = data.get('repository_id')
    metadata_field_id = data.get('metadata_field_id')
    checkbox_checked = data.get('checkbox_checked')
    
    valid_repo, repo_info = check_valid_repo(concentriq_service, repository_id)
    if not valid_repo:
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    try:
        metadata_field_id = int(metadata_field_id)
    except Exception as e:
        flash(f"{metadata_field_id} Not convertable to integer. ({e})", "error")
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    # Update ViewOnly status according to checkbox
    res = update_imageset_properties(concentriq_service.client, repository_id, metadata_field_id, viewOnly=False, study=checkbox_checked)
    
    if res is None:
        print(f"Error for field (id: {metadata_field_id} for imageset (id: {repository_id})")
        flash(f"Error for field (id: {metadata_field_id} for imageset (id: {repository_id})", "error")
    else:
        print(f"Change to Study for: field (id: {metadata_field_id})")

    # Get metadata fields and filter by content_type
    metadata_fields = get_metadata_fields(concentriq_service, repository_id)
    
    return render_template(
        'orderviewonly_page.html',
        repository_text=get_repository_text(repo_info),
        repository_id=repository_id,
        metadata_fields=metadata_fields
    )

@orderviewonly.route("/orderviewonly/OrderTemplateUpdate", methods=["POST"])
def OrderTemplateUpdate():
    # If not logged, return to login page
    concentriq_service = GLOBAL_STATE['service']
    if concentriq_service is None:
        return redirect(url_for("login.index"))
    
    repository_id = request.form.get("hiddenrepoid3")

    order_field_ids = request.form.get("templates_field_ids")
    valid_repo, repo_info = check_valid_repo(concentriq_service, repository_id)
    if not valid_repo:
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    try:
        order_field_ids = [int(num.strip()) for num in order_field_ids.split(',')]
    except Exception as e:
        flash(f"{order_field_ids} Not convertable to integer. ({e})", "error")
        return render_template(
            'orderviewonly_page.html',
            repository_id=repository_id,
        )
    
    metadata_fields = get_metadata_fields(concentriq_service, repository_id)
    filtered_metadata_fields = [metadata_field for metadata_field in metadata_fields if metadata_field.get('content_type') == 'image']
    
    # Update order of fields as the templates
    filters = {
        'imageSetId': [repository_id]
    }
    res_ = concentriq_service.client.get_metadata_fields(filters)
    
    metadata_fields_repos = [field_.get('id') for field_ in res_.get('fields')]

    order = 1
    for fields_id_ in order_field_ids:
        if fields_id_ in metadata_fields_repos:
            res = update_imageset_properties(concentriq_service.client, repository_id, fields_id_, orderNumber=order)
            if res is None:
                print(f"Error for field (id: {fields_id_} for imageset (id: {repository_id})")
                flash(f"Error for field (id: {fields_id_} for imageset (id: {repository_id})", "error")
            else:
                print(f"For field (id: {fields_id_}) set order: {order}")
                order+=1
        else:
            print(f"Skip field: {fields_id_}")
            flash(f"Skip field: {fields_id_}", "success")
    
    for filtered_metadata_field in filtered_metadata_fields:
        if filtered_metadata_field.get('id') not in order_field_ids:
            new_order = filtered_metadata_field.get('order')+order
            res = update_imageset_properties(concentriq_service.client, repository_id, filtered_metadata_field.get('id'), orderNumber=new_order)
            if res is None:
                print(f"Error for field (id: {filtered_metadata_field.get('id')} for imageset (id: {repository_id})")
                flash(f"Error for field (id: {filtered_metadata_field.get('id')} for imageset (id: {repository_id})", "error")
            else:
                print(f"For field (id: {filtered_metadata_field.get('id')}) set order: {new_order}")

    # Get metadata fields and filter by content_type
    metadata_fields = get_metadata_fields(concentriq_service, repository_id)
    filtered_metadata_fields = sorted(metadata_fields, key=lambda x: x['content_type'])
    
    return render_template(
        'orderviewonly_page.html',
        repository_text=get_repository_text(repo_info),
        repository_id=repository_id,
        metadata_fields=filtered_metadata_fields
    )

#####################################################################
############################# FUNCTIONS #############################
#####################################################################

def get_imageset_properties(concentriq, imageSetId:int):
    """Sends an XML file which is a representation of the annotations for a given image.
    Args:
        image_id (int): image unique ID.
    """
    response = concentriq.session.get('{}/metadata-fields/imageSetProperties/{}/'.format(concentriq.endpoint, imageSetId))
    return concentriq._handle_and_parse_response_text_data(response)

def update_imageset_properties(concentriq, imageSetId:int, field_id:int, orderNumber:int=None, viewOnly:bool=None, study:bool=None):
    data = {}
    if orderNumber is not None:
        data['orderNumber']=orderNumber
    if viewOnly is not None:
        data['viewOnly']=viewOnly
    if study is not None:
        data['study']=study
    res = concentriq.session.patch(
        '{}/metadata-fields/{}/imageSetProperties/{}/'.format(concentriq.endpoint, field_id, imageSetId),
        json=data)
    return concentriq._handle_and_parse_response_text_data(res)

def get_metadata_fields(concentriq_service, repository_id):
    metadata_fields = []
    filters = {
        'imageSetId': [repository_id]
    }
    res1 = concentriq_service.client.get_metadata_fields(filters)
    res2 = get_imageset_properties(concentriq_service.client, repository_id)
    
    for field1 in res1.get('fields'):
        for field2 in res2.get('properties'):
            if field2.get('fieldId') == field1.get('id'):
                metadata_field = {
                    'id': field1.get('id'),
                    'name': field1.get('name'),
                    'content_type': field1.get('resourceType'),
                    'order': field2.get('orderNumber'),
                    'viewonly': field2.get('viewOnly'),
                    'study': field2.get('study')
                }
                metadata_fields.append(metadata_field)
                
    return metadata_fields

def check_valid_repo(concentriq_service, repository_id):
    try:
        repository_id = int(repository_id)
    except Exception as e:
        flash(f"{repository_id} Not convertable to integer. ({e})", "error")
        return False, None
    
    repository_info = concentriq_service.client.get_imageset(repository_id)
    if repository_info is None:
        flash(f"Issue with this repository id: {repository_id}. Please check on Concentriq if this repository exist.", "error")
        return False, None
    
    return True, repository_info

def get_repository_text(repository_info):
    repository_name = repository_info.get('name')
    repository_id = repository_info.get('id')
    return f'for repository "{repository_name}" with id: {repository_id}'