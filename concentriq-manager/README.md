# How to access to concentriq service in the code

Using following part in python file you can access to the concentriq_service get at login.

```
from app.config import GLOBAL_STATE
concentriq_service = GLOBAL_STATE['service']
```

# add new app

In `main_page.html`, add following part with adapted `app_main_page` and `app_name`.

```
<div class="base-button">
    <button class="base-button" onclick="window.location.href='/app_main_page';">
        app_name
    </button>
</div>
```

In `run.py`, add following part with adapted `app_route_location` and `app_route_name`.

```
from app.routes.app_route_location import app_route_name
app.register_blueprint(app_route_name)
```
