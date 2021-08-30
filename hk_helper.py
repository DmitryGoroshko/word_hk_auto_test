import pywinauto
import time


# function for clearing current hotkey combination in listBox
def clear_current_hotkey(current_hk_listBox, word_keyboard_set_window_wr):
    while current_hk_listBox.texts():
        current_hk_listBox.set_focus()
        current_hk_listBox.click_input()
        if word_keyboard_set_window_wr.is_active():
            current_hk_listBox.type_keys("{UP}{ENTER}")


""" def print_list(input_list, err): function for printing result list
# second param. 1 - only Error
# second param. 0 - all"""


def print_list(input_list, err):
    for a in input_list:
        if ((err == 1) and (a[2] == 'Error')) or (err == 0):
            print(a)


# decode function: does the "B" sector exist
def decode_b(code_str):
    if code_str[4] == '1':
        return int(1)
    else:
        return int(0)


def add_vk_alt(name: str) -> str:
    return '{VK_MENU down}' + name + '{VK_MENU up}'


def add_vk_ctrl(name: str) -> str:
    return '{VK_CONTROL down}' + name + '{VK_CONTROL up}'


def add_vk_shift(name: str) -> str:
    return '{VK_SHIFT down}' + name + '{VK_SHIFT up}'


def add_gui_alt(name: str) -> str:
    return name + 'Alt+'


def add_gui_ctrl(name: str) -> str:
    return name + 'Ctrl+'


def add_gui_shift(name: str) -> str:
    return name + 'Shift+'


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
        if len(a_list) > 2:
            prefix_gui_list.extend([prefix_gui + str(a_list[2])])
    else:
        prefix_gui_list.extend([prefix_gui + str(a_list[1])])

    # Shift postfix
    if code_str[3] == '1':
        postfix_1 = '{VK_SHIFT down}'
        postfix_2 = '{VK_SHIFT up}'
        postfix_gui_list.extend([',Shift+' + str(b_list[1])])
        if len(b_list) > 2:
            postfix_gui_list.extend([',' + str(b_list[2])])
    elif code_str[4] == '1':
        postfix_gui_list.extend([',' + str(b_list[1])])

    if len(postfix_gui_list):
        for i in prefix_gui_list:
            for j in postfix_gui_list:
                gui_list.extend([i + j])
    else:
        gui_list.extend(prefix_gui_list)

    # b-segment (second ABC block)
    if code_str[4] == '0':
        return [str(prefix_1 + a_list[0] + prefix_2), gui_list]
    else:
        return [str(prefix_1 + a_list[0] + prefix_2 + postfix_1 + b_list[0] + postfix_2), gui_list]


# typing hotkey combinations
def type_hk(input_list, hk_edit):
    hk_edit.set_focus()
    hk_edit.click_input()
    pywinauto.keyboard.send_keys(input_list[0])


def read_hk(new_hk_edit, apply_button, word_current_hotkey_listBox):
    # combination recognized:
    if len(new_hk_edit.texts()) and (new_hk_edit.texts() != ' '):
        apply_button.click()
        recog_text = word_current_hotkey_listBox.texts()
        if len(recog_text):
            return recog_text[0][0]
        else:
            return 'Not recognized'
    else:
        return 'Not recognized'


# reading the entered combination and checking the correctness
def check_hk(input_list, recog_hk_str):
    typed_hk_str = input_list[1][0]
    status_hk_str = "Error"

    if recog_hk_str != "Not recognized":
        # checking start: typed_hk_str and recog_hk_str
        for i in input_list[1]:
            if recog_hk_str == i:
                status_hk_str = "OK"
                break
        # recognized empty line
    else:
        recog_hk_str = "Not recognized"
        status_hk_str = "Error"
    return [typed_hk_str, recog_hk_str, status_hk_str]


def check_focus(word_main_window, word_keyboard_set_window_wr, word_keyboard_set_window):
    # if the focus out of window
    if not word_keyboard_set_window_wr.is_active():
        time.sleep(1)
        word_main_window.minimize()
        word_main_window.maximize()
        word_main_window.set_focus()
        word_keyboard_set_window_wr.set_focus()
        word_keyboard_set_window.wait('visible')

        if not word_keyboard_set_window_wr.is_active():
            print("!!!Critical Error. Window was inactive!!!")
            exit(404)


def convert_input_data(input_data: list):
    result = []
    if len(input_data) == 0:
        return result
    else:
        for i_input in input_data:
            for j_alphabet in alphabet_list:
                if i_input == j_alphabet[1]:
                    result.extend([j_alphabet])
        return result


""" abc_list: alphabet for hotkey combinations
first pos - VK, second - visualisation in Word, third - shift visualisation "Shift + 1" or "!" for ex."""
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
            ['2', '2', '@'],
            ['3', '3', '#'],
            ['4', '4', '$'],
            ['5', '5', '{%}'],
            ['6', '6', '{^}'],
            ['7', '7', '&'],
            ['8', '8', '*'],
            ['9', '9', '('],
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

symbols_list = [[',', ',', '<'],
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

spec_symbols_alphabet = [['{BACKSPACE}', 'Backspace'],
                         ['{SPACE}', 'Space'],
                         ['{ENTER}', 'Return'],
                         ['{VK_UP}', 'Up'],
                         ['{VK_RIGHT}', 'Right'],
                         ['{DOWN}', 'Down'],
                         ['{LEFT}', 'Left'],
                         ['{HOME}', 'Home'],
                         ['{END}', 'End'],
                         ['{PGUP}', 'Page Up'],
                         ['{PGDN}', 'Page Down'],
                         ['{INS}', 'Insert'],
                         ['{DEL}', 'Del'],
                         ['{SCROLLLOCK}', 'Scroll Lock'],
                         ['{VK_LWIN}', 'WinKey']]

# commented because the alphabet is too large and it will take about 60 hours to run through all combination
alphabet_list = abc_list
alphabet_list.extend(num_list)
alphabet_list.extend(spec_symbols_alphabet)
alphabet_list.extend(func_list)
alphabet_list.extend(symbols_list)
alphabet_list.extend(num_pad_list)

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
