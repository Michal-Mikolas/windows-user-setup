import time
from windows.windows import Windows
from windows.settings import Settings
import config

windows = Windows(config.cache_path)
# settings = Settings(config.cache_path)

windows.clear_desktop(keep=['acrobat', 'chrome', 'backup', 'tor', 'nn.xlsx', 'ED7BA470-8E54-465E-825C-99712043E01C', 'vp2', 'moba'])
windows.setup_taskbar_search('icon')
windows.setup_taskbar_news('off')
windows.unpin_from_taskbar(['edge', 'store'])
windows.pin_to_taskbar(['chrome', 'adobe acrobat', 'outlook', 'microsoft word', 'excel'])

windows.show_desktop_this_pc()
