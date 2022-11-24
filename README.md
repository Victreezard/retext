# retext
Python script for OBS that resizes text frames of all slides in a PowerPoint presentation

# Getting Started
## Using the script
1. Make sure that PowerPoint presentation is open and enabled
2. Texts can be converted to all caps by entering their slide numbers in *Slide Points*
3. *Font Sizes* can be modified as well
4. Hit *Retext Slides* to start
![retext](https://user-images.githubusercontent.com/14161440/203749130-dbfd8621-13a4-45ea-8ccb-19a117397f03.gif)


## Rename slides saved as images for sorting purposes
1. Click *Rename*
2. Select the folder of images to be renamed
![retext_rename](https://user-images.githubusercontent.com/14161440/203753780-7c6e47ae-9b2d-4f83-9a61-9760f084237c.gif)


# Object References
- TextRange.ChangeCase - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchangecase
- Slide.Master - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.slide.master
- p_align - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.paragraphformat.alignment
- v_anchor - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.textframe.verticalanchor
- zorder - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.shape.zorder
