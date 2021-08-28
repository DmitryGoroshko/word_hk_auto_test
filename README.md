# Word hotkey auto test
Test task for auto test Word app for hot key module

## Python auto test program description
### About the program
This python script is autotest fot MS Word appliction hotkey module. It's generate all possiple hot key combinations, type and check recognition. 

### System requirements
My configuration:
    - Windows 10 64-bit
    - MS Office ProfessionalPlus 2019 Word 2102 build 13801.20864 32-bit
    - PyCharm 2021.2 (Community Edition 64-bit) 
    - Python 3.9.5

### Pre-conditions 
pywinauto lib
please start your default browser before starting the script (beause of the Shift+F1 hotkey)

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
3.	Сommands list GUI:
    - divided into two columns 
    - left column presents command names 
    - command names same as on the Ribbon
    - right column displays assigned hotkeys
    - organized as Ribbon tabs
4.	Commands group:
    - name same as Ribbon tab
    - ‘+’/’-’ Button located to the left of the group name
    - expanded with the ‘+’ Button
    - contracted with the ‘-’ Button
5.  Commands list scroll bar:
    - mouse click
    - mouse scroll
    - keyboard Up Down
    - keyboard PageUp PageDown
    - keyboard Home End
6.	Information about the selected command:
    - located under the commands list
    - includes the command name in bold
    - includes the command description
    - command description as in the command tooltip on the Ribbon
7.	Command selection:	    
    - Clicking will select a command
    - Hotkeys of the selected command can be modified
    - Backspace keyboard button remove all existing hotkeys
    - Pressing new keys combination on the keyboard will add the combination to the hotkeys list. (подробно реализовано автотестом)
    - Until the selection in the commands list is not changed, the last entered hotkey can be changed by entering another combination
    - Selecting another command in the list will store the entered hotkeys
    - If the combination is already used by another command, a message box will appear, asking if the combination should be reassigned.
8.  Closing Hotkeys window
    - ОК button
        - closes the dialog
        - applies all the changes
    - Cancel button
        - closes the dialog
        - does not make any change to commands hotkeys 
9.	After the hotkeys are changed using the dialog
    - corresponding Ribbon tooltips are updated to reflect the changed hotkeys
    - New hotkeys commands must be saved
    - New hotkeys commands recorded in the registry

## Test cases
**Pre-conditions:**
1.	OS required for the tested application.
2.	Launch the application.

### Test case #1.1 Hotkeys dialog modal (scene closed)
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	make sure that the scene closed | scene closed |
|2.	if not – close it | - |
|3.	open Hotkeys dialog | Hotkeys dialog opened |
|4.	make actions to parent application components (out from Hotkeys window, but inside parent application area) | - |
|4.1.	mouse click | application does not respond to actions |
|4.2.	mouse double click | application does not respond to actions |
|4.3.	mouse drug and drop test file from explorer to tested application | application does not respond to actions |
|4.4.	keyboard input [a…z][1…0] | application does not respond to actions |
|4.5.	type hotkey (Ctrl+E; Ctrl+А; Ctrl+5; Ctrl+6; Ctrl+0) | application does not respond to actions |

### Test case #1.2 Hotkeys dialog modal (scene opened)
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	open the scene | scene opened |
|2.	open Hotkeys dialog | Hotkeys dialog opened |
|3.	make actions to scene (out from Hotkeys window, but inside parent application area) | scene are not enable does not respond to actions |
|3.1.	mouse click | application does not respond to actions |
|3.2.	mouse double click | application does not respond to actions |
|3.3.	mouse drug and drop test file from explorer to tested application | application does not respond to actions |
|3.4.	keyboard input [a…z][1…0] | application does not respond to actions |
|3.5.	type hotkey (Ctrl+E; Ctrl+А; Ctrl+5; Ctrl+6; Ctrl+0) | application does not respond to actions |

### Test case #2 GUI elements
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog| Hotkeys dialog opened |
|2.	The title of the window dialog | "Hotkeys" |
|3.	Does commands list exist? | yes |
|4.	Does information area exist? | yes |
|5.	Does button "OK" exist? | yes |
|6.	Does button "Cancel" exist? | yes |
|7. Resize window by mouse at the bottom right window corner | the window has not been resized |

### Test case #3 Commands list
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                          | Hotkeys dialog opened |
|2. Is Commands list divided into two columns?   | yes |
|3. Does the left column represent command names?| yes |
|4. Does the right column displays assigned hotkeys? | yes |
|5. Is command names in the commands list same as on the Ribbon | yes | 
|6. Is command group names in the commands list same as the Ribbon tabs | yes |

### Test case #3 Commands list
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                          | Hotkeys dialog opened |
|2. Is Commands list divided into two columns?   | yes |
|3. Does the left column represent command names?| yes |
|4. Does the right column displays assigned hotkeys? | yes |
|5. Is command names in the commands list same as on the Ribbon | yes | 
|6. Is command group names in the commands list same as the Ribbon tabs | yes |

### Test case #4 Commands group
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                          | Hotkeys dialog opened |
|2. Is ‘+’/’-’ Button located to the left of the group name | yes |
|3. Press '-' Button to the left of the group name          |  Command groups contracted |
|4. Is the '-' Button changed to '+' Button?    | yes |
|5. Press '+' Button to the left of the group name | Command groups expanded |
|6. Is the '+' Button changed to '-' Button?    | yes |

### Test case #5 Commands list scroll bar
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog   | Hotkeys dialog opened |
|2. Move to Commands list scroll bar | - |
|3. Move mouse coursor to the "V" button | - |
|4. Click to the "V" button | Commands list scrolling down |
|5. Move mouse coursor to the "^" button | - |
|6. Click to the "^" button | Commands list scrolling up |
|7. Move mouse coursor to the scroll bar button | - |
|8. Press left mouse and move down | Commands list scrolling down |
|9. Move mouse up | Commands list scrolling up |
|10. release the left mouse button | - |
|11. Set focus to the commands list | - |
|12. scroll down with the mouse wheel | Commands list scrolling down |
|13. scroll up with the mouse wheel | Commands list scrolling up |
|14. Set focus to the commands list | - |
|15. press and release the down key on the keyboard | Commands list scrolling down |
|16. press and release the up key on the keyboard | Commands list scrolling up |
|17. press and release the PgDn key on the keyboard | Commands list scrolling down |
|18. press and release the PgUp on the keyboard | Commands list scrolling up |
|19. press and release the End key on the keyboard | Commands list scrolling to bottom |
|20. press and release the PgUp on the keyboard | Commands list scrolling to top |

### Test case #6 Information about the selected command
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog   | Hotkeys dialog opened |
|2. Is information area located under the commands list | yes |
|3. Click to the command name in comands list | the command cell is highlighted, Information area includes the command name in **bold** and the command description |
|4. Is command description as in the command tooltip on the Ribbon | yes |

### Test case #7.1 Command selection, add new hotkey combination
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog   | Hotkeys dialog opened |
|2. Click to the command name without hot key combination in commands list | the command cell is highlighted |
|3. Enter all allpossible hotkey combination | hotkey combination added to the right column in front of command name. Recognized key combination must match with the entered from the keyboard| 
|4. Select another command in the commansd list | last entered hotkey saved |
|5. Select previous command again | the command cell is highlighted |
|6. Enter new hotkey combination | hotkey combination added to the right column in front of command name |
|7. Select another command in the commansd list | last entered hotkey added to the hotkey combination. Two hotkey combination for the one command |
|8. Select previous command again | the command cell is highlighted |
|9. Press Backspace keyboard button | all hotkey combination cleared |

### Test case #7.2 Message box "combination is already used by another command"
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog   | Hotkeys dialog opened |
|2. Click to the command name without hot key combination in commands list | the command cell is highlighted |
|3. Enter new hotkey combination | hotkey combination added to the right column in front of command name |
|4. Select another command in the commansd list | last entered hotkey saved |
|5. Enter the same hotkey combination | a message box will appear, asking if the combination should be reassigned |
|6. Click to the "Cancel" button in message box | the hotkey combination not saved |
|7. Enter the same hotkey combination | a message box will appear, asking if the combination should be reassigned |
|8. Click to the "OK" button in message box | the hotkey combination saved |

### Test case #8.1 Closing Hotkeys window ОК button
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog   | Hotkeys dialog opened |
|2. Click to the command (remember it) name without hot key combination in commands list | the command cell is highlighted |
|3. Enter new hotkey combination | hotkey combination added to the right column in front of command name |
|4. Select another command in the commansd list | last entered hotkey saved |
|5. Click "OK" button in Hotkeys window | Hotkeys dialog closed|
|6. Select the Ribon in the main window | Ribon selected |
|7. Find command from step 2 on the Ribbon, set mouse focus | new hotkey from step 3 in Ribbon tooltips displayed | 
|8. Enter hotkey from step 3 in the main window | command from step 2 was called |
|9. Find hotkey from step 3 in the registry | new hotkey from step 3 saved in registry |

### Test case #8.2 Closing Hotkeys window Cancel button
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog   | Hotkeys dialog opened |
|2. Click to the command (remember it) name without hot key combination in commands list | the command cell is highlighted |
|3. Enter new hotkey combination | hotkey combination added to the right column in front of command name |
|4. Select another command in the commansd list | last entered hotkey saved |
|5. Click "Close" button in Hotkeys window | Hotkeys dialog closed |
|6. Select the Ribon in the main window | Ribon selected |
|7. Find command from step 2 on the Ribbon, set mouse focus | new hotkey from step 3 in Ribbon tooltips not displayed | 
|8. Enter hotkey from step 3 in the main window | command from step 2 was not called |
|9. Find hotkey from step 3 in the registry | new hotkey from step 3 was not saved in registry |
