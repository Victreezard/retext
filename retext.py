from os import listdir, rename
from os.path import abspath, join, splitext
from re import compile
from time import perf_counter
from tkinter import filedialog
import config
import obspython as obs
import tkinter as tk
import win32com.client


# ---------------Custom Functions---------------

class Retext():
    # Defaults: Align Center, Anchor Middle, ZOrder BringToFront
    def __init__(self, width=0, height=0, left=0, top=0, font_size=50, point_font_size=50,
                 p_align=2, v_anchor=3, zorder=2):
        try:
            self.ppt = win32com.client.GetActiveObject(config.POWERPOINT).ActivePresentation
        except Exception as e:
                # print(e)
                print("ERROR: A PowerPoint Presentation must be opened and editable")
                self.ppt = None
                return

        self.width = self.ppt.Slides(1).Master.Width if width <= 0 else width
        self.height = (self.ppt.Slides(1).Master.Height /
                       3) if height <= 0 else height
        self.left = left
        self.top = (self.height * 2) if top <= 0 else top
        self.font_size = font_size
        self.point_font_size = point_font_size
        self.p_align = p_align
        self.v_anchor = v_anchor
        self.zorder = zorder

    @staticmethod
    def get_number(text):
        """
        Return an integer from a text
        """
        return int(text.strip())

    @staticmethod
    def get_numbers(text):
        """
        Return a list of integers from a text
        """
        return list(map(int, text.strip().split()))

    def retext(self, points):
        """For each slide, reformat shapes that have non-empty text."""
        for slide in self.ppt.Slides:
            for shape in slide.Shapes:
                if shape.TextFrame.HasText and shape.TextFrame.TextRange != '':
                    # Slides with Titles will have different properties
                    if slide.SlideNumber in points:
                        case = config.UPPERCASE
                        font_size = self.point_font_size
                    else:
                        case = config.TITLECASE
                        font_size = self.font_size

                    shape.ZOrder(self.zorder)
                    shape.Height = self.height
                    shape.Top = self.top
                    shape.Left = self.left
                    shape.Width = self.width
                    shape.TextFrame.TextRange.ChangeCase(case)
                    shape.TextFrame.TextRange.Font.Size = font_size
                    shape.TextFrame.TextRange.ParagraphFormat.Alignment = self.p_align
                    shape.TextFrame.VerticalAnchor = self.v_anchor

                    # Unsure why but some slides needs these twice to get proper top value
                    shape.Top = self.top
                    shape.Height = self.height

def retext_callback(props, prop):
    time_start = perf_counter()

    global points_text
    global points_size_text
    global others_size_text

    rt = Retext()
    if rt.ppt is None: return

    if points_size_text:
        rt.point_font_size = rt.get_number(points_size_text)
    if others_size_text:
        rt.font_size = rt.get_number(others_size_text)
    # Parse title slide numbers if input is not empty
    points_text = rt.get_numbers(points_text) if points_text else []

    # print(f"{points_text=!s}")
    # print(f"{points_size_text=!s}")
    # print(f"{others_size_text=!s}")
    rt.retext(points_text)

    time_stop = perf_counter()
    time_elapsed = str(time_stop-time_start)
    obs.obs_data_set_string(settings_copy, "info_text",
                            formatted_time := f"{time_elapsed=!s}")
    print(formatted_time)

    return True

def browse_directory() -> str:
    root = tk.Tk()
    root.withdraw()

    dir_path = filedialog.askdirectory()
    if not dir_path:
        print("No folder was selected")
        return
    # print(f"{dir_path=!s}")
    return dir_path


def rename_callback(props, prop):
    """
    Remove leading letters and convert the remaining number to its equivalent alphabet.
    Return a list of unchanged filenames if any.
    PowerPoint has the filename format "Slide1" when saving slides as images.
    """
    dir_path = browse_directory()
    if not dir_path: return

    dir_path = abspath(dir_path)
    # print(f"{dir_path=!s}")

    letters = compile(r'^[a-zA-Z_]+')
    errors = []

    for filename in listdir(dir_path):
        try:
            name, ext = splitext(letters.sub('', filename))
            name = chr(ord('@') + int(name))
            rename(join(dir_path, filename), join(dir_path, name + ext))
        except Exception:
            errors.append(filename)
    if errors: print(f"Error occured for the following: {errors}")


# ---------------Script Global Functions---------------


def script_description():
    return """Retext"""


def script_defaults(settings):
    """Sets the default values of the Script's properties when 'Defaults' button is pressed"""
    pass


def script_update(settings):
    global points_text
    points_text = obs.obs_data_get_string(settings, "points_text")
    global points_size_text
    points_size_text = obs.obs_data_get_string(settings, "points_size_text")
    global others_size_text
    others_size_text = obs.obs_data_get_string(settings, "others_size_text")
    global settings_copy
    settings_copy = settings


def script_properties():
    props = obs.obs_properties_create()
    # Properties for slides containing important points
    obs.obs_properties_add_text(props, "points_info_text", "Points", obs.OBS_TEXT_INFO)
    obs.obs_properties_add_text(props, "points_text", "Slide Points", obs.OBS_TEXT_DEFAULT)
    obs.obs_properties_add_text(props, "points_size_text", "Font Size", obs.OBS_TEXT_DEFAULT)

    # Properties for other slides
    obs.obs_properties_add_text(props, "others_info_text", "Others", obs.OBS_TEXT_INFO)
    obs.obs_properties_add_text(props, "others_size_text", "Font Size", obs.OBS_TEXT_DEFAULT)

    # Action Buttons
    obs.obs_properties_add_button(props, "retext_button", "Retext Slides", retext_callback)
    obs.obs_properties_add_button(props, "rename_button", "Rename Saved PNGs", rename_callback)

    obs.obs_properties_add_text(props, "info_text", "", obs.OBS_TEXT_INFO)

    return props
