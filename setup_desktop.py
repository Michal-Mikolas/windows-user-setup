from windows.windows import Windows
from windows.settings import Settings
import config

windows = Windows(config.cache_path)
settings = Settings(config.cache_path)

windows.settings_language()
