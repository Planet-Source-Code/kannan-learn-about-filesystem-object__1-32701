
ReadMe: 

How this program works.

The mainform (frmMain) has a picture. Its corners were covered by 4 labels. Their property is set to transparent
during run time. 

To move the form during runtime Sendmessage is used. I think the releaseCapture is not needed.

Whatever the content you are viewing in the UI that is from the file "FileSystemDemo.txt". Its file format
is described in the file itself.

All the required information like, The topic information, whether it has any help associated and its 
description will be found in the file itself.

Read the file and fill the array with the function ReadAndFillArray. Then something is dynamic. Load the n number of labels 
equal to the n number of topics found in the file.

All the labels will have the caption as the topic title. On selecting the particular topic (nothing but label) the 
appropriate topic will be displayed in the right side pane (a text box. Can be changed to some other control). Just for a visual 
effect a green line will move along with the selected topic.

If the topic has any help available, View sample button will be visible. Click the button will show the form with the approp topic.
This is done by setting the topic index to the frmSample.  

Each and every topic will be displayed in a different frame. Form height and width will be adjusted based on the 
frame height and width. 

Please leave your valueable comments. That will help me to improve my skills. 

Thank you..
..kannan