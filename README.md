# word_hk_auto_test
test task for auto test Word app for hot key module

## Сheck list:
1.  Is Hotkeys a dialog modal:
    - dialog window is displayed
    - blocks user interaction with the parent application
    - not allow making any modification to a scene 
    - not allow making any modification to other application components 
    - can be run even if no scene is opened 
2.	Hotkeys dialog GUI elements:
    - commands list exists
    - information area exists
    - button OK exists
    - button Cancel exists
3.	Сommands list:
    - divided into two columns 
    - left column presents command names 
    - command names same as on the Ribbon
    - right column displays assigned hotkeys
    - organized as Ribbon tabs
4.	Commands group
    - name same as Ribbon tab
    - ‘+’/’-’ Button located to the left of the group name
    - expanded with the ‘+’ Button
    - contracted with the ‘-’ Button
    - clicking will select a command.
4.	Information about the selected command
    - located under the commands list
    - includes the command name in bold
    - includes the command description
    - command description as in the command tooltip on the Ribbon
5.	Hotkeys of the selected command can be modified
    - Backspace keyboard button will remove all existing hotkeys
    - Pressing new keys combination on the keyboard will add the combination to the hotkeys list. (подробно реализовано автотестом)
    - Until the selection in the commands list is not changed, the last entered hotkey can be changed by entering another combination
    - Selecting another command in the list will store the entered hotkeys
6.	If the combination is already used by another command, a message box will appear, asking if the combination should be reassigned.
    - ОК button
        - closes the dialog
        - applies all the changes
    - Cancel button
        - closes the dialog
        - does not make any change to commands hotkeys 
7.	After the hotkeys are changed using the dialog
    - corresponding Ribbon tooltips are updated to reflect the changed hotkeys
    - New hotkeys commands must be saved
    - New hotkeys commands recorded in the registry

## Case test

| Syntax | Description |
| ----------- | ----------- |
| Header | Title |
| Paragraph | Text |
| Paragraph | Text |
| Paragraph | Text |
| Paragraph | Text |
| Paragraph | Text |
