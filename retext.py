import PySimpleGUI as sg
import win32com.client


class Slide_Format():
    # Defaults: Align Center, Anchor Middle, ZOrder BringToFront
    def __init__(self, width=0, height=0, left=0, top=0, font_size=50, p_align=2, v_anchor=3,
                 zorder=2):
        self.ppt = win32com.client.GetActiveObject(
            "PowerPoint.Application").ActivePresentation

        self.width = self.ppt.Slides(1).Master.Width if width <= 0 else width
        self.height = (self.ppt.Slides(1).Master.Height / 3) if height <= 0 else height
        self.left = left
        self.top = (self.height * 2) if top <=0 else top
        self.font_size = font_size
        self.p_align = p_align
        self.v_anchor = v_anchor
        self.zorder = zorder

    def reformat(self, point_list):
        """For each slide, reformat shapes that have non-empty text."""
        for slide in self.ppt.Slides:
            for shape in slide.Shapes:
                if shape.TextFrame.HasText and shape.TextFrame.TextRange != '':
                    # Specified Slides will become Upper Case else Title Case
                    if slide.SlideNumber in point_list:
                        shape.TextFrame.TextRange.ChangeCase(3)
                    else:
                        shape.TextFrame.TextRange.ChangeCase(4)

                    shape.Width = self.width
                    shape.Height = self.height
                    shape.Left = self.left
                    shape.Top = self.top
                    shape.TextFrame.TextRange.Font.Size = self.font_size
                    shape.TextFrame.TextRange.ParagraphFormat.Alignment = self.p_align
                    shape.TextFrame.VerticalAnchor = self.v_anchor
                    shape.ZOrder(self.zorder)
            sg.one_line_progress_meter('Retexting', slide.SlideNumber, self.ppt.Slides.Count)

sg.theme('Black')
submit_button = 'Retext'
points_input = 'Points'
font_size = 'FontSize'

layout = [
    [sg.Text('Font Size')],
    [sg.InputText(k=font_size)],
    [sg.Text('Slides to Capitalize')],
    [sg.InputText(k=points_input)],
    [sg.Button(submit_button)]
]

slides = Slide_Format()
window = sg.Window('Retext', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    else:
        if values[font_size]:
            slides.font_size = int(values[font_size].strip())
        # Convert each list item from str to int
        point_list = list(map(int, values[points_input].strip().split()))

        slides.reformat(point_list)

window.close()
