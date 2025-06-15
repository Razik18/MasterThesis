from app.config import GLOBAL_STATE

import os
from flask import (
    Blueprint, render_template, request, redirect, url_for, flash, render_template, jsonify
)
import win32cred
from ConcentriqSDK.concentriq_service import ConcentriqService

login = Blueprint('login', __name__)

@login.route("/login")
def index():
    password, email="",""
    try:
        # Get credential from windows credential with name "ConcentriqProd"
        email_win, password_win = get_credentials()
        if email_win:
            email = email_win
        if password_win:
            password = password_win
    except:
        pass

    try:
        # Try to get from os environment variable password and mail
        if not password:
            password = os.environ.get('CONCENTRIQ_PASSWORD', '')
        if not email:
            email = os.environ.get('CONCENTRIQ_EMAIL', '')
    except:
        pass
    
    return render_template(
        'login.html',
        email = email,
        password = password,
    )

@login.route("/get_login", methods=["POST"])
def get_login():
    email = request.form.get("email", "").strip()
    password = request.form.get("password", "").strip()
    endpoint = request.form.get("concentriq_env", "").strip()

    if not email or not password or not endpoint:
        flash("Missing required settings (email, password, endpoint).", "error")
        return redirect(url_for("login.index"))

    try:
        concentriq_service = ConcentriqService(
            endpoint=endpoint,
            email=email, 
            password=password,
            pac_url=GLOBAL_STATE.get('pac_url'))
        
    except Exception as e:
        err_str = str(e)
        if '"status":401' in err_str or "login_failed" in err_str:
            flash("The login information is incorrect (401).", "error")
        else:
            flash(f"Error creating client: {e}", "error")
        return redirect(url_for("login.index"))

    if concentriq_service.is_logged():
        GLOBAL_STATE['service'] = concentriq_service
        print("ConcentriqService stored in GLOBAL_STATE:", GLOBAL_STATE['service'])
    else:
        print("Login failed. Service not stored.")
        return redirect(url_for("login.index"))

    flash("Login successfully.", "success")
    return redirect(url_for("main.index"))

@login.route('/get_env_options', methods=['GET'])
def get_env_options():
    """Allows to get different options for environment from config."""
    env_dict = {}
    try:
        for env_name in GLOBAL_STATE.get('endpoints'):
            env_dict[env_name] = GLOBAL_STATE.get('endpoints').get(env_name)
        return jsonify(env_dict)
    except Exception as e:
        return jsonify({"error": "Failed to retrieve environment options"}), 500

def get_credentials():
    try:
        cred = win32cred.CredRead("ConcentriqProd", win32cred.CRED_TYPE_GENERIC)
        username = cred['UserName']
        password = cred['CredentialBlob'].decode('utf-8').replace('\x00', '').strip()
        return username, password
    except Exception as e:
        print(f"Error retrieving credentials: {e}")
        return "", ""