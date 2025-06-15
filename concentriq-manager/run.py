import time
import threading
import webbrowser
from flask import (
    Flask
)
from flask_session import Session

from app.routes.login_routes import login
from app.routes.main_routes import main
from app.routes.ocr_routes import ocr
from app.routes.orderviewonly import orderviewonly

# -------------- Flask App Setup --------------
app = Flask(__name__, template_folder='app/templates', static_folder="app/static")
app.register_blueprint(login)
app.register_blueprint(main)
app.register_blueprint(ocr)
app.register_blueprint(orderviewonly)
app.config["SESSION_TYPE"] = "filesystem"
Session(app)

# -------------- Main --------------
def open_browser():
  time.sleep(1.25)
  webbrowser.open("http://127.0.0.1:5012/login")

if __name__ == "__main__":
    threading.Thread(target=open_browser).start()
    app.run(debug=False, port=5012)
