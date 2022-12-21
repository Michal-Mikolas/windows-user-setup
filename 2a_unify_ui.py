import time
from windows.windows import Windows
from windows.settings import Settings
import config

windows = Windows(config.cache_path)

# 1) Unify Windows UI
# windows.disable_transparency()
# windows.activate_dark_theme()
# windows.setup_languages('en-US,cs-CZ')
# windows.set_display_language('en-US')
windows.unify_ui()
