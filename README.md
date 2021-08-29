# Word hotkey auto test
Test task: 'Create auto test using Python for 1 test case. Use for example ‘Word’ application to test your autotest versus.' 

This is auto test for Word application (hot key module).  

## Python auto test program description
### About the program
This python script is autotest fot MS Word appliction hotkey module. 
It's generate all possiple hot key combinations, type and check recognition. 

### System requirements
My configuration:
- Windows 10 64-bit RUS
- MS Office ProfessionalPlus 2019 Word 2102 build 13801.20864 32-bit RUS
- PyCharm 2021.2 (Community Edition 64-bit) 
- Python 3.9.5

### Pre-conditions 
- pywinauto lib
- please start your default browser before starting the script (beause of the Shift+F1 hotkey)
- make sure you have the English keyboard layout

### How it works?
I did a little research on keyboard hotkey for the Word application.
As far as I was able to find out, this application accepts the following combinations as hot keys:

| # | Alt | Ctrl | Shift | Shift | B | Code|
| --- | --- | --- | --- | --- | --- | --- |
| 1  | 1 |   |   |   |   | **10000**|
| 2  | 1 |   |   |   | 1 | **10001**|
| 3  | 1 |   |   | 1 | 1 | **10011**|
| 4  | 1 |   | 1 |   |   | **10100**|
| 5  | 1 |   | 1 |   | 1 | **10101**|
| 6  | 1 |   | 1 | 1 | 1 | **10111**|
| 7  |   | 1 |   |   |   | **01000**|
| 8  |   | 1 |   |   | 1 | **01001**|
| 9  |   | 1 |   | 1 | 1 | **01011**|
| 10 |   | 1 | 1 |   |   | **01100**|
| 11 |   | 1 | 1 |   | 1 | **01101**|
| 12 |   | 1 | 1 | 1 | 1 | **01111**|
| 13 | 1 | 1 |   |   |   | **11000**|
| 14 | 1 | 1 |   |   | 1 | **11001**|
| 15 | 1 | 1 |   | 1 | 1 | **11011**|
| 16 | 1 | 1 | 1 |   |   | **11100**|
| 17 | 1 | 1 | 1 |   | 1 | **11101**|
| 18 | 1 | 1 | 1 | 1 | 1 | **11111**|
| 19 |   |   | 1 |   |   | **00100**|
| 20 |   |   | 1 |   | 1 | **00101**|
| 21 |   |   | 1 | 1 | 1 | **00111**|

===============================================================================

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
|1.	make sure that the scene closed, if not – close it                                                          | scene closed                            |
|2.	open Hotkeys dialog                                                                                         | Hotkeys dialog window opened            |
|3.	make actions to parent application components (out from Hotkeys window, but inside parent application area) | -                                       |
|3.1.	mouse click                                                                                             | application does not respond to actions |
|3.2.	mouse double click                                                                                      | application does not respond to actions |
|3.3.	mouse drug and drop test file from explorer to tested application                                       | application does not respond to actions |
|3.4.	keyboard input [a…z][1…0]                                                                               | application does not respond to actions |
|3.5.	type hotkey (Ctrl+E; Ctrl+А; Ctrl+5; Ctrl+6; Ctrl+0)                                                    | application does not respond to actions |

### Test case #1.2 Hotkeys dialog modal (scene opened)
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	open the scene                                                                      | scene opened |
|2.	open Hotkeys dialog                                                                 | Hotkeys dialog window opened |
|3.	make actions to scene (out from Hotkeys window, but inside parent application area) | scene are not enable does not respond to actions |
|3.1.	mouse click                                                                     | application does not respond to actions |
|3.2.	mouse double click                                                              | application does not respond to actions |
|3.3.	mouse drug and drop test file from explorer to tested application               | application does not respond to actions |
|3.4.	keyboard input [a…z][1…0]                                                       | application does not respond to actions |
|3.5.	type hotkey (Ctrl+E; Ctrl+А; Ctrl+5; Ctrl+6; Ctrl+0)                            | application does not respond to actions |

### Test case #2 GUI elements
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                         | Hotkeys dialog window opened |
|2.	The title of the window dialog                              | "Hotkeys" |
|3.	Check if commands list exists                               | commands list exists |
|4.	Check if information area exists                            | information area exists  |
|5.	Check if button "OK" exists                                 | button "OK" exists |
|6.	Check if button "Cancel" exists                             | button "Cancel" exists |
|7. Resize window by mouse at the bottom right window corner    | the window has not been resized |

### Test case #3 Commands list
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                                         | Hotkeys dialog window opened |
|2. Check if Commands list divided into two columns                             | Commands list divided into two columns |
|3. Check if left column represent command names                                | left column represent command names |
|4. Check if right column displays assigned hotkeys                             | right column displays assigned hotkeys |
|5. Check if command names in the Commands list same as on the Ribbon           | command names in the Commands list same as on the Ribbon | 
|6. Check if command group names in the Commands list same as the Ribbon tabs   | command group names in the Commands list same as the Ribbon tabs |

### Test case #4 Commands group
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                             | Hotkeys dialog window opened |
|2. Check if ‘+’/’-’ Button located to the left of the group name   | ‘+’/’-’ Button located to the left of the group name |
|3. Press '-' Button to the left of the group name                  | Command groups contracted |
|4. Check if '-' Button changed to '+' Button                       | '-' Button changed to '+' Button |
|5. Press '+' Button to the left of the group name                  | Command groups expanded |
|6. Check if '+' Button changed to '-' Button                       | '+' Button changed to '-' Button |

### Test case #5 Commands list scroll bar
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                                 | Hotkeys dialog window opened |
|2. Move mouse coursor to Commands list scroll bar                      | mouse coursor is on the Commands list scroll bar |
|3. Move mouse coursor to the "V" button                                | mouse coursor is on the "V" button  |
|4. Click to the "V" button                                             | Commands list scrolling down |
|5. Move mouse coursor to the "^" button                                | mouse coursor is on the "^" button |
|6. Click to the "^" button                                             | Commands list scrolling up |
|7. Move mouse coursor to the scroll bar element                        | mouse coursor to the scroll bar element |
|8. Press left mouse and move down                                      | Commands list scrolling down |
|9. Move mouse up                                                       | Commands list scrolling up |
|10. release the left mouse button                                      | nothing happened |
|11. Set focus to the commands list                                     | the command cell is highlighted |
|12. scroll down with the mouse wheel                                   | Commands list scrolling down |
|13. scroll up with the mouse wheel                                     | Commands list scrolling up |
|14. Set focus to the commands list                                     | the command cell is highlighted |
|15. press and release the "Down" key on the keyboard                   | Commands list scrolling down |
|16. press and release the "Up" key on the keyboard                     | Commands list scrolling up |
|17. press and release the "PgDn" key on the keyboard                   | Commands list scrolling down |
|18. press and release the "PgUp" on the keyboard                       | Commands list scrolling up |
|19. press and release the "End" key on the keyboard                    | Commands list scrolling to bottom |
|20. press and release the "Home' on the keyboard                       | Commands list scrolling to top |

### Test case #6 Information about the selected command
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                                 | Hotkeys dialog window opened |
|2. Check if information area located under the commands list           | nformation area located under the commands list |
|3. Click to the command name in comands list                           | the command cell is highlighted | 
|4. Check if Information area includes the command name in **bold**     | Information area includes the command name in **bold** |
|5. Check if Information area includes the command  description         | Information area includes the command  description |
|6. Check if command description as in the command tooltip on the Ribbon| command description as in the command tooltip on the Ribbon |

### Test case #7.1 Command selection, add and save new hotkey, clear current hotkey
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                                                 | Hotkeys dialog window opened |
|2. Click to the command name were hot key combination cell in commands list is empty   | the command cell is highlighted |
|3. Enter one of hotkey combination "Ctrl" + one of [a...z]                             | hotkey combination added to the right column in front of command name |
|4. Check if recognized key combination match with the typed from the keyboard          | Recognized key combination match with the typed from the keyboard |
|5. Select another command in the commansd list                                         | last entered hotkey saved |
|6. Select previous command again                                                       | the command cell is highlighted |
|7. Enter new hotkey combination                                                        | hotkey combination added to the right column in front of command name |
|8. Select another command in the commansd list                                         | last entered hotkey added to the hotkey combination and saved |
|9. Check if two hotkey combination are for the same command                            | two hotkey combination are for the same command |
|10. Select previous command in the commansd list again                                 | the command cell is highlighted |
|10. Press "Backspace" keyboard button                                                  | all hotkey combination for this command cleared |

### Test case #7.2 Command list add new correct hotkey 
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                                                 | Hotkeys dialog window opened |
|2. Click to the command name were hot key combination cell in commands list is empty   | the command cell is highlighted |
|3. Enter hotkey combination "Ctrl" + "Shift" + "A"                                     | hotkey combination added to the right column in front of command name |
|4. Check if recognized key combination match with the typed from the keyboard          | Recognized key combination match with the typed from the keyboard "Ctrl+Shift+A" |
|5. Select another command in the commansd list                                         | last entered hotkey saved |
|6. Select previous command again                                                       | the command cell is highlighted |
|7. Enter hotkey combination "Ctrl" + "Shift" + "Z"                                     | hotkey combination added to the right column in front of command name |
|8. Check if recognized key combination match with the typed from the keyboard          | Recognized key combination match with the typed from the keyboard "Ctrl+Shift+Z"|
|9. Select another command in the commansd list                                         | last entered hotkey added to the hotkey combination and saved |
|10. Check if two hotkey combination are for the same command                           | two hotkey combination are for the same command "Ctrl+Shift+A" and "Ctrl+Shift+Z" |
|11. Select previous command in the commansd list again                                 | the command cell is highlighted |
|12. Press "Backspace" keyboard button                                                  | all hotkey combination for this command cleared |

### Test case #7.3 Command list add already used hotkey
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog   | Hotkeys dialog window opened |
|2. Click to the command name without hot key combination in commands list | the command cell is highlighted |
|3. Enter new hotkey combination | hotkey combination added to the right column in front of command name |
|4. Select another command in the commansd list | last entered hotkey saved |
|5. Enter the same hotkey combination | a message box will appear, asking if the combination should be reassigned |
|6. Click to the "Cancel" button in message box | the hotkey combination not saved |
|7. Enter the same hotkey combination | a message box will appear, asking if the combination should be reassigned |
|8. Click to the "OK" button in message box | the hotkey combination saved |

### Test case #7.3 Command list try to add system hotkey 
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                                                 | Hotkeys dialog window opened |
|2. Click to the command name were hot key combination cell in commands list is empty   | the command cell is highlighted |
|3. Enter hotkey combination "Ctrl" + "Esc"                                             | hotkey combination not added. Windows Star menu opened. Focus changed |
|4. Click to the command name were hot key combination cell in commands list is empty   | the command cell is highlighted |
|5. Enter hotkey combination "Alt" + "Space"                                            | hotkey combination not added. Standard Windows context menu of the dialog opened |

### Test case #7.4 Command list try to add incorrect hotkey 
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                                                 | Hotkeys dialog window opened |
|2. Click to the command name were hot key combination cell in commands list is empty   | the command cell is highlighted |
|3. Enter hotkey combination "E" + "4"                                                  | hotkey combination not recognized, not added |
|4. Click to the command name were hot key combination cell in commands list is empty   | the command cell is highlighted |
|5. Enter hotkey combination "P" + "-"                                                  | hotkey combination not recognized, not added |

### Test case #7.5 Command list try to add correct hotkey advanced combinations 
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog                                                                 | Hotkeys dialog window opened |
|2. Click to the command name were hot key combination cell in commands list is empty   | the command cell is highlighted |
|3. Enter hotkey combination "E" + "4"                                                  | hotkey combination not recognized, not added |











### Test case #8.1 Closing Hotkeys window ОК button
| Test Steps | Expected Result |
| ----------- | ----------- |
|1.	Open Hotkeys dialog   | Hotkeys dialog window opened |
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
|1.	Open Hotkeys dialog   | Hotkeys dialog window opened |
|2. Click to the command (remember it) name without hot key combination in commands list | the command cell is highlighted |
|3. Enter new hotkey combination | hotkey combination added to the right column in front of command name |
|4. Select another command in the commansd list | last entered hotkey saved |
|5. Click "Close" button in Hotkeys window | Hotkeys dialog closed |
|6. Select the Ribon in the main window | Ribon selected |
|7. Find command from step 2 on the Ribbon, set mouse focus | new hotkey from step 3 in Ribbon tooltips not displayed | 
|8. Enter hotkey from step 3 in the main window | command from step 2 was not called |
|9. Find hotkey from step 3 in the registry | new hotkey from step 3 was not saved in registry |
