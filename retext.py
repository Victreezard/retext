import cfg
import win32com.client
import PySimpleGUI as sg


class Retext():
    # Defaults: Align Center, Anchor Middle, ZOrder BringToFront
    def __init__(self, width=0, height=0, left=0, top=0, font_size=50, title_font_size=50,
                 p_align=2, v_anchor=3, zorder=2):
        self.ppt = win32com.client.GetActiveObject(
            cfg.POWERPOINT).ActivePresentation

        self.width = self.ppt.Slides(1).Master.Width if width <= 0 else width
        self.height = (self.ppt.Slides(1).Master.Height /
                       3) if height <= 0 else height
        self.left = left
        self.top = (self.height * 2) if top <= 0 else top
        self.font_size = font_size
        self.title_font_size = title_font_size
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

    def reformat(self, title_list):
        """For each slide, reformat shapes that have non-empty text."""
        for slide in self.ppt.Slides:
            for shape in slide.Shapes:
                if shape.TextFrame.HasText and shape.TextFrame.TextRange != '':
                    # Slides with Titles will have different properties
                    if slide.SlideNumber in title_list:
                        case = cfg.UPPERCASE
                        font_size = self.title_font_size
                    else:
                        case = cfg.TITLECASE
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

            sg.one_line_progress_meter(
                'Retexting', slide.SlideNumber, self.ppt.Slides.Count)


sg.theme('Black')
SUBMIT_BUTTON = 'Retext'
TITLES_INPUT = 'Slide Numbers'
FONT_SIZE_INPUT = 'Font Size'
TITLE_FONT_SIZE_INPUT = 'Title Font Size'

layout = [
    [sg.Fr('Titles', [
        [sg.Text(TITLES_INPUT)],
        [sg.InputText(k=TITLES_INPUT)],
        [sg.Text(FONT_SIZE_INPUT)],
        [sg.InputText(k=TITLE_FONT_SIZE_INPUT)],
    ])],
    [sg.Fr('Others', [
        [sg.Text(FONT_SIZE_INPUT)],
        [sg.InputText(k=FONT_SIZE_INPUT)]
    ])],
    [sg.Button(SUBMIT_BUTTON)]
]

rt = Retext()
window = sg.Window('Reformat slides', layout, icon=cfg.ICON_PATH)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    else:
        if values[FONT_SIZE_INPUT]:
            rt.font_size = rt.get_number(values[FONT_SIZE_INPUT])
        if values[TITLE_FONT_SIZE_INPUT]:
            rt.title_font_size = rt.get_number(values[TITLE_FONT_SIZE_INPUT])
        # Parse title slide numbers if input is not empty
        title_list = rt.get_numbers(
            values[TITLES_INPUT]) if values[TITLES_INPUT] else []

        rt.reformat(title_list)

window.close()
