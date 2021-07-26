from os.path import abspath, join
import sys


def resolve_path(relative_path):
    """
    Get absolute path to resource.
    Reference: https://stackoverflow.com/a/13790741
    """
    try:
        # PyInstaller creates a temp folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = abspath('.')
    return join(base_path, relative_path)

RESOURCES = 'resources'  # directory name
ICON = 'tlag.ico'
ICON_PATH = resolve_path(join(RESOURCES, ICON))
POWERPOINT = "PowerPoint.Application"
UPPERCASE = 3
TITLECASE = 4

class Tooltip():
    RENAME_FOLDER_BROWSER = 'Browse the folder of slides saved as images and rename for sorting purposes.'
