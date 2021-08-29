import time
import win32api
import pywinauto
from pywinauto.application import Application

win32api.LoadKeyboardLayout("00000409", 1)  # ENG
start_time = time.time()

# launch the Word application

app = Application(backend="uia").start("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
                                       timeout=3)
# Word main window
word_main_window = app.Word
word_main_window.wait('visible')

# navigation on "File" tab
pywinauto.keyboard.send_keys("{TAB 5}{UP}{ENTER}")

# Word settings window
word_parametr_window = word_main_window.ПараметрыWord
word_parametr_window.wait('visible')

# navigation "Word settings -> Ribbon settings -> Settings "
pywinauto.keyboard.send_keys("{DOWN 7}{TAB 3}{ENTER}")
word_keyboard_set_window = word_parametr_window.НастройкаКлавиатуры
word_keyboard_set_window.wait('visible')
word_keyboard_set_window_wr = word_keyboard_set_window.wrapper_object()

# current hotkey ListBox object and wrapper
word_current_hotkey_listBox = word_keyboard_set_window.ListBox3.wrapper_object()

# new hot key Edit object and wrapper
new_hk_edit = word_keyboard_set_window.Edit2.wrapper_object()
new_hk_edit.draw_outline()

# "Apply" Button wrapper
apply_button = word_keyboard_set_window.НазначитьButton.wrapper_object()

# alphabet for hotkey combinations
# first pos - VK, second - visualisation in Word, third - shift visualisation "Shift + 1" or "!" for ex.
abc_list = [['A', 'A'],
            ['B', 'B'],
            ['C', 'C'],
            ['D', 'D'],
            ['E', 'E'],
            ['F', 'F'],
            ['G', 'G'],
            ['H', 'H'],
            ['I', 'I'],
            ['J', 'J'],
            ['K', 'K'],
            ['L', 'L'],
            ['M', 'M'],
            ['N', 'N'],
            ['O', 'O'],
            ['P', 'P'],
            ['Q', 'Q'],
            ['R', 'R'],
            ['S', 'S'],
            ['T', 'T'],
            ['U', 'U'],
            ['V', 'V'],
            ['W', 'W'],
            ['X', 'X'],
            ['Y', 'Y'],
            ['Z', 'Z']]

num_list = [['1', '1', '!'],
            #            ['2', '2', '@'],
            #            ['3', '3', '#'],
            #            ['4', '4', '$'],
            #            ['5', '5', '{%}'],
            #            ['6', '6', '{^}'],
            #            ['7', '7', '&'],
            #            ['8', '8', '*'],
            #            ['9', '9', '('],
            ['0', '0', ')']]

func_list = [['{F1}', 'F1'],
             ['{F2}', 'F2'],
             ['{F3}', 'F3'],
             ['{F4}', 'F4'],
             ['{F5}', 'F5'],
             ['{F6}', 'F6'],
             ['{F7}', 'F7'],
             ['{F8}', 'F8'],
             ['{F9}', 'F9'],
             ['{F10}', 'F10'],
             ['{F11}', 'F11'],
             ['{F12}', 'F12']]

num_pad_list = [['{VK_NUMPAD0}', 'Num 0'],
                ['{VK_NUMPAD1}', 'Num 1'],
                ['{VK_NUMPAD2}', 'Num 2'],
                ['{VK_NUMPAD3}', 'Num 3'],
                ['{VK_NUMPAD4}', 'Num 4'],
                ['{VK_NUMPAD5}', 'Num 5'],
                ['{VK_NUMPAD6}', 'Num 6'],
                ['{VK_NUMPAD7}', 'Num 7'],
                ['{VK_NUMPAD8}', 'Num 8'],
                ['{VK_NUMPAD9}', 'Num 9'],
                ['{VK_ADD}', 'Num +'],
                ['{VK_DECIMAL}', 'Num .'],
                ['{VK_DIVIDE}', 'Num /'],
                ['{VK_MULTIPLY}', 'Num *']]

simbols_list = [[',', ',', '<'],
                ['.', '.', '>'],
                ['/', '/', '?'],
                [';', ';', ':'],
                ["'", "'", '"'],
                ['[', '[', '{'],
                [']', ']', '}'],
                ['\\', '\\', '|'],
                ['-', '-', '_'],
                ['=', '=', '+'],
                ['`', '`', '{~}']]

spec_simb_list = [['{BACKSPACE}', 'Backspace'],
                  ['{SPACE}', 'Пробел'],
                  ['{ENTER}', 'Return'],
                  ['{VK_UP}', 'Стрелка вверх'],
                  ['{VK_RIGHT}', 'Стрелка вправо'],
                  ['{DOWN}', 'Стрелка вниз'],
                  ['{LEFT}', 'Стрелка влево'],
                  ['{HOME}', 'Home'],
                  ['{END}', 'End'],
                  ['{PGUP}', 'PgUp'],
                  ['{PGDN}', 'PgDn'],
                  ['{INS}', 'Ins'],
                  ['{DEL}', 'Del'],
                  ['{SCROLLLOCK}', 'Scroll Lock'],
                  ['{VK_LWIN}', 'WinKey']]

# code for all possible combination Alt+Ctrl+Shift+a_ABC,Shift+b_ABC
code_list = ['10000',
             '10001',
             '10011',
             '10100',
             '10101',
             '10111',
             '01000',
             '01001',
             '01011',
             '01100',
             '01101',
             '01111',
             '11000',
             '11001',
             '11011',
             '11100',
             '11101',
             '11111',
             '00100',
             '00101',
             '00111']

# commented because the alphabet is too large and it will take about 60 hours to run through all combination
# alphavit_list = abc_list
# alphavit_list.extend(num_list)
# alphavit_list.extend(spec_simb_list)
# alphavit_list.extend(func_list)
# alphavit_list.extend(simbols_list)
# alphavit_list.extend(num_pad_list)
alphavit_list = num_list


result_list = [["Typed", "Recognized", "Result"]]


# function for clearing current hotkey combination
def clear_current_hotkey():
    for a in range(len(word_current_hotkey_listBox.texts())):
        word_current_hotkey_listBox.set_focus()
        word_current_hotkey_listBox.click_input()
        word_current_hotkey_listBox.type_keys("{UP}{ENTER}")


# function for printing result list
# second param. 1 - only Error
# second param. 0 - all
def print_list(input_list, err):
    for a in range(len(input_list)):
        if ((err == 1) and (input_list[a][2] == 'Error')) or (err == 0):
            print(input_list[a])


# typing hotkey combinations
def type_hk(input_list):
    new_hk_edit.set_focus()
    new_hk_edit.click_input()
    pywinauto.keyboard.send_keys(input_list[0])


# advanced function typing hotkey combinations
def advanced_type_hk(input_list):
    type_hk(input_list)

    # if the focus out of window
    if not word_keyboard_set_window_wr.is_active():
        word_keyboard_set_window_wr.set_focus()
        if not word_keyboard_set_window_wr.is_active():
            print("!!!Critical Error. Window was inactive!!!")
            print_list(result_list, 0)
            exit(404)
        # try again
        type_hk(input_list)


# reading the entered combination and checking the correctness
def read_and_check_hk(input_list):
    typed_hk_str = input_list[1][0]
    status_hk_str = "Error"

    # combination recognized:
    if len(new_hk_edit.texts()) and (new_hk_edit.texts() != ' '):
        apply_button.click()
        current_hk_list = word_current_hotkey_listBox.texts()
        # recognized something
        if len(current_hk_list):
            # recognized hotkey str
            recog_hk_str = str(current_hk_list[0][0])
            # checking start: typed_hk_str and recog_hk_str
            for i in range(len(input_list[1])):
                if recog_hk_str == input_list[1][i]:
                    status_hk_str = "OK"
                    break
                else:
                    status_hk_str = "Error"
            # checking end: typed_hk_str and recog_hk_str
            clear_current_hotkey()
        # recognized empty line
        else:
            recog_hk_str = "Not recognized"
            status_hk_str = "Error"
    # combination not recognized:
    else:
        recog_hk_str = "Not captured"
        status_hk_str = "Error"
    return [typed_hk_str, recog_hk_str, status_hk_str]


# decode function: does the "B" sector exist
def decode_b(code_str):
    if code_str[4] == '1':
        return int(1)
    else:
        return int(0)


# decode Alt Ctrl Shift combination alphabet
def decode_ctrl(code_str, a_list, b_list):
    prefix_gui = ''
    prefix_1 = ''
    prefix_2 = ''
    postfix_1 = ''
    postfix_2 = ''
    prefix_gui_list = []
    postfix_gui_list = []
    gui_list = []

    # Alt
    if code_str[0] == '1':
        prefix_1 += '{VK_MENU down}'
        prefix_2 += '{VK_MENU up}'
        prefix_gui += 'Alt+'
    # Ctrl
    if code_str[1] == '1':
        prefix_1 += '{VK_CONTROL down}'
        prefix_2 += '{VK_CONTROL up}'
        prefix_gui += 'Ctrl+'
    # Shift prefix
    if code_str[2] == '1':
        prefix_1 += '{VK_SHIFT down}'
        prefix_2 += '{VK_SHIFT up}'
        prefix_gui_list.extend([prefix_gui + 'Shift+' + str(a_list[1])])
        prefix_gui_list.extend([prefix_gui + str(a_list[2])])
    else:
        prefix_gui_list.extend([prefix_gui + str(a_list[1])])

    # Shift postfix
    if code_str[3] == '1':
        postfix_1 = '{VK_SHIFT down}'
        postfix_2 = '{VK_SHIFT up}'
        postfix_gui_list.extend([',Shift+' + str(b_list[1])])
        postfix_gui_list.extend([',' + str(b_list[2])])
    elif code_str[4] == '1':
        postfix_gui_list.extend([',' + str(b_list[1])])

    if len(postfix_gui_list):
        for i in range(len(prefix_gui_list)):
            for j in range(len(postfix_gui_list)):
                gui_list.extend([prefix_gui_list[i] + postfix_gui_list[j]])
    else:
        gui_list.extend(prefix_gui_list)

    # b-segment (second ABC block)
    if code_str[4] == '0':
        return [str(prefix_1 + a_list[0] + prefix_2), gui_list]
    else:
        return [str(prefix_1 + a_list[0] + prefix_2 + postfix_1 + b_list[0] + postfix_2), gui_list]


# current command may have assigned hotkeys
clear_current_hotkey()

# main loop of the program, run through all possible key combinations
for i_ctrl in range(len(code_list)):
    # does b ABC block exist
    if decode_b(code_list[i_ctrl]) == 0:
        b_abc_block_len = 1
    else:
        b_abc_block_len = len(alphavit_list)

    # a ABC block
    for i_alph in range(len(alphavit_list)):

        # b ABC block
        for j_alph in range(b_abc_block_len):
            current_comb_list = decode_ctrl(code_list[i_ctrl], alphavit_list[i_alph], alphavit_list[j_alph])
            # print(current_comb_list)

            # typing hotkey
            advanced_type_hk(current_comb_list)
            # read, recognition, check
            temp_result_list = read_and_check_hk(current_comb_list)
            # adding results
            result_list.extend([temp_result_list])
# print results
print("Tested in: " + str(time.time() - start_time) + " sec")
print("Total checked: " + str(len(result_list) - 1) + " combinations")

print_list(result_list, 1)
app.kill()

"""
Shift + F1 - web help
Ctrl + Shift + Win, Shift + Win - not displayed
Ctrl + Shift + drop-down menu - not displayed
"""
