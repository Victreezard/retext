import PySimpleGUI as sg
import win32com.client


ppt = win32com.client.GetActiveObject(
    "PowerPoint.Application").ActivePresentation


class MySlide():
    def __init__(self, width, height, left, top, font_size, p_align):
        self.width = width
        self.height = height
        self.left = left
        self.top = top
        self.font_size = font_size
        self.p_align = p_align


def reshape(point, subtitle, point_list):
    for count_slide in range(1, ppt.Slides.Count + 1):
        shape_config = point if str(count_slide) in point_list else subtitle
        for count_shape in range(1, ppt.Slides(count_slide).Shapes.Count + 1):
            current_shape = ppt.Slides(count_slide).Shapes(count_shape)
            if current_shape.TextFrame.HasText:
                current_shape.Width = shape_config.width
                current_shape.Height = shape_config.height
                current_shape.Left = shape_config.left
                current_shape.Top = shape_config.top
                current_shape.TextFrame.TextRange.Font.Size = shape_config.font_size
                current_shape.TextFrame.TextRange.ParagraphFormat.Alignment = shape_config.p_align
                current_shape.TextFrame.VerticalAnchor = 3
                current_shape.ZOrder(2)


sg.theme('DarkBlue')
submit_button = 'GO!'
exit_button = 'Exit'
points_input = 'Points'
point_size = 'PointSize'
subtitle_size = 'SubtitleSize'

layout = [
    [sg.Text('Enter Point Font Size'), sg.InputText(k=point_size)],
    [sg.Text('Enter Subtitle Font Size'), sg.InputText(k=subtitle_size)],
    [sg.HorizontalSeparator()],
    [sg.Text('Enter POINT SLIDE NUMBERS separated by spaces'),
     sg.InputText(k=points_input)],
    [sg.Button(submit_button), sg.Button(exit_button)]
]

my_point = MySlide(700, 800, 10, 0, 95, 1)
my_subtitle = MySlide(1450, 150, 0, 650, 60, 2)
window = sg.Window('Reformat Slide Text', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == exit_button:
        break
    else:
        if values[point_size]:
            my_point.font_size = int(values[point_size])
        if values[subtitle_size]:
            my_subtitle.font_size = int(values[subtitle_size])
        point_list = values[points_input].split()
        reshape(my_point, my_subtitle, point_list)

window.close()
