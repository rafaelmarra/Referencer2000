# Referencer2000
VBA script for CorelDraw, draws lines and text boxes for referencing drawing components

This is a simple script that I've made to make my life easier. I use it to make numbered references in drawings using CorelDraw, as I am currently working in a pattent office and sometimes we do that kind of thing.

Just import the files into your Gloal Variables and set a hotkey to the DrawLineAtClick() command. When the hotkey is pressed, the user needs to click on the screen twice. Those 2 clicks will be the start and finish point of the line segment.
After that, a pop up window with a number field will show up. Whatever reference number (or text) you want your pointer to have will be added here. Press the button and you are done.

The code takes into account the line orientation for the placing of the text, in order to avoid clipping. Also, if the number field is left blank, no text box will be created.
