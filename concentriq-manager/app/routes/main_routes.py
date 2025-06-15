from app.config import GLOBAL_STATE

from flask import (
    Blueprint, render_template, redirect, url_for, render_template
)

main = Blueprint('main', __name__)

@main.route("/main")
def index():
    concentriq_service = GLOBAL_STATE['service']
    if concentriq_service is None:
        return redirect(url_for("login.index"))
    return render_template(
        'main_page.html'
    )