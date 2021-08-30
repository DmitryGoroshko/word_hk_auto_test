import time
import win32api
import pywinauto
from pywinauto.application import Application
import hk_helper


def word_hotkey_auto_test(input_data: list):
    # launch the Word application
    app = Application(backend="uia").start("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
                                           timeout=3)
    # Word main window
    word_main_window = app.Word
    # print(type(word_main_window))
    word_main_window.wait('visible')

    # navigation on "File" tab
    pywinauto.keyboard.send_keys("{TAB 5}{UP}{ENTER}")

    # Word Options window
    word_parameter_window = word_main_window.WordOptions
    word_parameter_window.wait('visible')

    # navigation "Word Options -> Customize Ribbon -> Customize "
    pywinauto.keyboard.send_keys("{DOWN 7}{TAB 3}{ENTER}")
    word_keyboard_set_window = word_parameter_window.CustomizeKeyboard

    word_keyboard_set_window.wait('visible')
    word_keyboard_set_window_wr = word_keyboard_set_window.wrapper_object()

    # current hotkey ListBox object and wrapper
    word_current_hotkey_listBox = word_keyboard_set_window.ListBox3.wrapper_object()

    # new hot key Edit object and wrapper
    new_hk_edit = word_keyboard_set_window.Edit2.wrapper_object()
    new_hk_edit.draw_outline()

    # "Apply" Button wrapper. Customize Keyboard window
    apply_button = word_keyboard_set_window.AssignButton.wrapper_object()

    result_list = [["Typed", "Recognized", "Result"]]

    # converting a sequence of keys to VK and Shift visualisation mode
    input_seq = hk_helper.convert_input_data(input_data)

    # current command may have assigned hotkeys
    hk_helper.clear_current_hotkey(word_current_hotkey_listBox, word_keyboard_set_window_wr)

    # main loop of the program, run through all possible key combinations
    for i_ctrl in hk_helper.code_list:

        # does b ABC block exist
        if hk_helper.decode_b(i_ctrl) == 0:
            b_abc_block_len = 1
        else:
            b_abc_block_len = len(input_seq)

        # a ABC block
        for i_alph in input_seq:
            # b ABC block
            for j_alph in range(b_abc_block_len):
                current_comb_list = hk_helper.decode_ctrl(i_ctrl, i_alph, input_seq[j_alph])

                # typing hotkey
                hk_helper.check_focus(word_main_window, word_keyboard_set_window_wr, word_keyboard_set_window)
                hk_helper.type_hk(current_comb_list, new_hk_edit)
                hk_helper.check_focus(word_main_window, word_keyboard_set_window_wr, word_keyboard_set_window)

                # read typed into Edit combination
                recog_str = hk_helper.read_hk(new_hk_edit, apply_button, word_current_hotkey_listBox)
                # check if typed and recognized hot key the same
                temp_result_list = hk_helper.check_hk(current_comb_list, recog_str)
                # clear current hot key editBox
                hk_helper.clear_current_hotkey(word_current_hotkey_listBox, word_keyboard_set_window_wr)
                # adding results to output list
                result_list.extend([temp_result_list])

    # kill Word app
    app.kill()
    return result_list


if __name__ == '__main__':
    # preparation
    win32api.LoadKeyboardLayout("00000409", 1)  # ENG
    start_time = time.time()

    # type test keyset here
    special_set_list = ['A', 'Z', '1', '0', 'F1', 'F12', 'Num 0', 'Num 9', '-', '.', 'Up', 'Del', 'Backspace']

    # run autotet
    result = word_hotkey_auto_test(special_set_list)

    # print results
    print("Tested in: " + str(time.time() - start_time) + " sec")
    print("Total checked: " + str(len(result) - 1) + " combinations")
    print("Errors report:")
    hk_helper.print_list(result, 1)
