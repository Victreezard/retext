# retext
Python script for OBS that resizes text frames of all slides in a PowerPoint presentation

![RetetxtScript](https://user-images.githubusercontent.com/14161440/203720232-985e6378-00a2-40d6-86fc-79f4fd14fe08.png)

# Getting Started
## 
TODO add GIF
- Before
![Presentation_Before](https://user-images.githubusercontent.com/14161440/203720435-6c13890f-dba3-420a-9e6f-8f42604dc5be.png)
- After
![Presentation_After](https://user-images.githubusercontent.com/14161440/203720545-2b9727e1-68bf-42ac-92ae-06674b61034b.png)

## Rename slides saved as images for sorting purposes.
- Before
![Before_Rename](https://user-images.githubusercontent.com/14161440/203721222-228cb1fa-64af-456d-911e-d7f0eb4af214.png)
- After
![After_Rename](https://user-images.githubusercontent.com/14161440/203721250-4e107e0e-c229-40db-83bd-6cf4b88fb6e2.png)


Object References:
- TextRange.ChangeCase - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchangecase
- Slide.Master - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.slide.master
- p_align - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.paragraphformat.alignment
- v_anchor - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.textframe.verticalanchor
- zorder - https://docs.microsoft.com/en-us/office/vba/api/powerpoint.shape.zorder
