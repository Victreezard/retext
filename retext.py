import PySimpleGUI as sg
import win32com.client


class Slide_Format():
    # Defaults: Align Center, Anchor Middle, ZOrder BringToFront
    def __init__(self, width=0, height=0, left=0, top=0, font_size=50, p_align=2, v_anchor=3,
                 zorder=2):
        self.ppt = win32com.client.GetActiveObject(
            "PowerPoint.Application").ActivePresentation

        self.width = self.ppt.Slides(1).Master.Width
        self.height = int(self.ppt.Slides(1).Master.Height * 0.3)
        self.left = 0
        self.top = self.ppt.Slides(1).Master.Height - self.height
        self.font_size = font_size
        self.p_align = p_align
        self.v_anchor = v_anchor
        self.zorder = zorder

    def reformat(self, point_list):
        """For each slide, change the properties of the first object that contains a text."""
        for slide in self.ppt.Slides:
            for shape in slide.Shapes:
                if shape.TextFrame.HasText:
                    # Slides marked as points will become Upper Case for emphasis else Title Case
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

# Set UI theme and control names
sg.theme('DarkBlue')
submit_button = 'GO!'
exit_button = 'Exit'
points_input = 'Points'
font_size = 'FontSize'

layout = [
    [sg.Text('Enter Title Font Size')],
    [sg.InputText(k=font_size)],
    [sg.HorizontalSeparator()],
    [sg.Text('Enter SLIDE NUMBERS of POINTS separated by spaces')],
    [sg.InputText(k=points_input)],
    [sg.Button(submit_button), sg.Button(exit_button)]
]

slides = Slide_Format()
window = sg.Window('Reformat Slide Text', layout)

# UI starts here
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == exit_button:
        break
    else:
        if values[font_size]:
            slides.font_size = int(values[font_size])
        point_list = list(map(int, values[points_input].split()))

        slides.reformat(point_list)

window.close()
