# word_hk_auto_test
test task for auto test Word app for hot key module

1. Item 1
1. Item 2
1. Item 3
   2. Item 3a
   2. Item 3b

Сheck list:

1.	Is Hotkeys a dialog modal:
   1.	dialog window is displayed
   1.	blocks user interaction with the parent application
   1.	not allow making any modification to a scene 
   1.	not allow making any modification to other application components 
   1.	can be run even if no scene is opened 
2.	Hotkeys dialog GUI elements:
a.	commands list exists
b.	information area exists
c.	button OK exists
d.	button Cancel exists
3.	Сommands list:
a.	divided into two columns 
b.	left column presents command names 
c.	command names same as on the Ribbon
d.	right column displays assigned hotkeys
e.	organized as Ribbon tabs
f.	commands group
i.	name same as Ribbon tab
ii.	‘+’/’-’ Button located to the left of the group name
iii.	expanded with the ‘+’ Button
iv.	contracted with the ‘-’ Button
g.	clicking will select a command.
4.	Information about the selected command
a.	located under the commands list
b.	includes the command name in bold
c.	includes the command description
d.	command description as in the command tooltip on the Ribbon
5.	Hotkeys of the selected command can be modified
a.	Backspace keyboard button will remove all existing hotkeys
b.	Pressing new keys combination on the keyboard will add the combination to the hotkeys list. (подробно реализовано автотестом)
c.	Until the selection in the commands list is not changed, the last entered hotkey can be changed by entering another combination
d.	Selecting another command in the list will store the entered hotkeys
6.	If the combination is already used by another command, a message box will appear, asking if the combination should be reassigned.
a.	ОК button
i.	closes the dialog
ii.	applies all the changes
b.	Cancel button
i.	closes the dialog
ii.	does not make any change to commands hotkeys 
7.	After the hotkeys are changed using the dialog
a.	corresponding Ribbon tooltips are updated to reflect the changed hotkeys
b.	New hotkeys commands must be saved
c.	New hotkeys commands recorded in the registry

