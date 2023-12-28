//highest error nr: BIC059
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Xml;
using WindowsInput.Native;

namespace Speech
{
    public partial class MainWindow
    {
        bool updating_bic_toggling_data_on = true;

        CollectionView cv_bic_off;
        CollectionView cv_bic_general;
        CollectionView cv_bic_mouse;
        CollectionView cv_bic_pressing;
        CollectionView cv_bic_inserting;
        CollectionView cv_bic_dict_always;
        CollectionView cv_bic_dict_better;

        public class BuiltInCommand
        {
            public string name;
            public string name_firstupper { get; set; } //first letter uppercase
            public string description { get; set; }
            public string key_combination { get; set; }
            public short max_executions { get; set; }
            public bool enabled { get; set; }

            public bic_type type; //bic = built-in command

            public string keys = null;

            public VirtualKeyCode vkc = VirtualKeyCode.VK_7;

            public bool use_contains = false; //instead of using ==

            public BuiltInCommand()
            {

            }

            public BuiltInCommand(string Description, bic_type Type, string Name, string Keys,
                string Key_combination, short Max_executions, bool Use_contains)
            {
                description = Description;
                type = Type;
                name = Name;
                keys = Keys;
                key_combination = Key_combination;
                max_executions = Max_executions;
                use_contains = Use_contains;

                enabled = true;
                name_firstupper = name.FirstCharToUpper();
            }

            public BuiltInCommand(string Description, bic_type Type, string Name, VirtualKeyCode Vkc,
                string Key_combination, short Max_executions)
            {
                description = Description;
                type = Type;
                name = Name;
                vkc = Vkc;
                key_combination = Key_combination;
                max_executions = Max_executions;

                enabled = true;
                name_firstupper = name.FirstCharToUpper();
            }

            public BuiltInCommand(string Description, bic_type Type, string Name, short Max_executions)
            {
                description = Description;
                type = Type;
                name = Name;
                max_executions = Max_executions;

                key_combination = "No";
                enabled = true;
                name_firstupper = name.FirstCharToUpper();
            }
        }

        List<BuiltInCommand> list_bic_off;
        List<BuiltInCommand> list_bic_general;
        List<BuiltInCommand> list_bic_mouse;
        List<BuiltInCommand> list_bic_general_and_mouse; //enabled only
        List<BuiltInCommand> list_bic_keys_pressing;
        List<BuiltInCommand> list_bic_char_inserting;
        List<BuiltInCommand> list_bic_dictation_always;
        List<BuiltInCommand> list_bic_dictation_better;

        void update_bic_grammar()
        {
            try
            {
                while (recognition_suspended)
                {
                    Thread.Sleep(10);
                }

                recognition_suspended = true;

                Mouse.OverrideCursor = Cursors.Wait;
                SW.Bmode.Visibility = Visibility.Hidden;

                //THRswitch_to.Abort(); //abort only causes problems
                thread_suspend1 = true;

                //wait for suspend:
                while (thread_suspended1 == false)
                {
                    Thread.Sleep(10);
                }

                unload_grammar_if_loaded(grammar_off_mode, grammar_off_mode_name);
                unload_grammar_if_loaded(grammar_builtin_commands, grammar_builtin_commands_name);
                unload_grammar_if_loaded(grammar_dictation_commands, grammar_dictation_commands_name);

                if (are_all_bic_off_disabled() == false)
                    grammar_off_mode = create_off_mode_grammar();
                else
                    grammar_off_mode = null;

                if (are_all_bic_general_and_mouse_disabled() == false)
                    grammar_builtin_commands = create_builtin_commands_grammar();
                else
                    grammar_builtin_commands = null;

                if (are_all_bic_dictation_disabled() == false)
                    grammar_dictation_commands = create_dictation_commands_grammar();
                else
                    grammar_dictation_commands = null;

                if (grammar_off_mode != null)
                    recognizer.LoadGrammar(grammar_off_mode);

                if (grammar_dictation_commands != null)
                    recognizer.LoadGrammar(grammar_dictation_commands);

                if (grammar_builtin_commands != null)
                    recognizer.LoadGrammar(grammar_builtin_commands);

                //debug only:
                //display_grammars_status();

                //loaded grammar is enabled by default
                if (current_mode == mode.off)
                {
                    toggle_grammar(true, grammar_type.grammar_off_mode);
                    toggle_grammar(false, grammar_type.grammar_builtin_commands);
                    toggle_grammar(false, grammar_type.grammar_dictation_commands);
                }
                else if (current_mode == mode.command)
                {
                    toggle_grammar(false, grammar_type.grammar_off_mode);
                    toggle_grammar(true, grammar_type.grammar_builtin_commands);
                    toggle_grammar(false, grammar_type.grammar_dictation_commands);
                }
                else if (current_mode == mode.dictation)
                {
                    toggle_grammar(false, grammar_type.grammar_off_mode);
                    toggle_grammar(false, grammar_type.grammar_builtin_commands);
                    toggle_grammar(true, grammar_type.grammar_dictation_commands);
                }

                //debug only:
                //MessageBox.Show(is_bic_in_general_and_mouse_enabled(bic_type.open_app).ToString());

                if (is_bic_in_general_and_mouse_enabled(bic_type.switch_to_app))
                {
                    apps_switching = true;

                    if (is_grammar_loaded(grammar_apps_switching_name))
                    {
                        if (current_mode == mode.command)
                            toggle_grammar(true, grammar_type.grammar_apps_switching);
                    }
                    else
                    {
                        if (grammar_apps_switching != null)
                        {
                            recognizer.LoadGrammar(grammar_apps_switching);

                            toggle_grammar(true, grammar_type.grammar_apps_switching);
                        }
                    }
                }
                else
                {
                    apps_switching = false;

                    unload_grammar_if_loaded(grammar_apps_switching, grammar_apps_switching_name);
                }

                if (is_bic_in_general_and_mouse_enabled(bic_type.open_app))
                {
                    apps_opening = true;

                    if (is_grammar_loaded(grammar_apps_opening_name))
                    {
                        if (current_mode == mode.command)
                            toggle_grammar(true, grammar_type.grammar_apps_opening);
                    }
                    else
                    {
                        if (grammar_apps_opening != null)
                        {
                            recognizer.LoadGrammar(grammar_apps_opening);

                            toggle_grammar(true, grammar_type.grammar_apps_opening);
                        }
                    }
                }
                else
                {
                    apps_opening = false;

                    unload_grammar_if_loaded(grammar_apps_opening, grammar_apps_opening_name);
                }

                thread_suspend1 = false;

                recognition_suspended = false;

                SW.Bmode.Visibility = Visibility.Visible;
                Mouse.OverrideCursor = null;

                //debug only
                //display_grammars_status();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void create_bic_lists() //bic = built-in command
        {
            try
            {
                list_bic_off = new List<BuiltInCommand>();

                list_bic_off.Add(new BuiltInCommand("Turn on speech recognition",
                    bic_type.turn_on, turn_on, 1));

                list_bic_general = new List<BuiltInCommand>();

                list_bic_general.Add(new BuiltInCommand("Turn off speech recognition",
                    bic_type.turn_off, turn_off, 1));
                list_bic_general.Add(new BuiltInCommand("Switch to dictation mode",
                    bic_type.switch_to_dictation, switch_to_dictation_mode, 1));
                list_bic_general.Add(new BuiltInCommand("Show speech recognition window",
                    bic_type.show_speech_recognition, show_speech_recognition, 1));
                list_bic_general.Add(new BuiltInCommand("Hide speech recognition window",
                    bic_type.hide_speech_recognition, hide_speech_recognition, 1));
                list_bic_general.Add(new BuiltInCommand("Open an application. You can use full\r\n" +
                    "application name or part of it.",
                    bic_type.open_app, open_app_str + " app_name", null, "No", 1, true));
                list_bic_general.Add(new BuiltInCommand("Switch to an open application. You can use\r\n" +
                    "full application window title or part of it.",
                    bic_type.switch_to_app, switch_to_app_str + " window_title", null, "No", 1, true));
                list_bic_general.Add(new BuiltInCommand("Minimize current application",
                    bic_type.minimize_that, "minimize", 1));
                list_bic_general.Add(new BuiltInCommand("Maximize current application",
                    bic_type.maximize_that, "maximize", 1));
                list_bic_general.Add(new BuiltInCommand("Restore current application",
                    bic_type.restore_that, "restore", 1));
                list_bic_general.Add(new BuiltInCommand("Close current application",
                    bic_type.close_that, "close that", 1));
                list_bic_general.Add(new BuiltInCommand("Start macro recording if recording window\r\n" +
                "is open.",
                    bic_type.start_recording, "start recording", 1));
                list_bic_general.Add(new BuiltInCommand("Stop macro recording if recording window\r\n" +
                "is open.",
                    bic_type.stop_recording, "stop recording", 1));
                list_bic_general.Add(new BuiltInCommand("Set X and Y in mouse action window\r\n" +
                "to current mouse cursor position.",
                    bic_type.get_position, "get position", 1));

                list_bic_mouse = new List<BuiltInCommand>();

                list_bic_mouse.Add(new BuiltInCommand("Show the mousegrid and move the mouse\r\n" +
                    "cursor to the chosen figure.",
                    bic_type.move, "move", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Show the mousegrid and left click\r\n" +
                    "in the chosen figure.",
                    bic_type.left, "west", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Show the mousegrid and right click\r\n" +
                    "in the chosen figure.",
                    bic_type.right, "east", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Show the mousegrid and double click\r\n" +
                    "in the chosen figure.",
                    bic_type.double2, "Double", null, "Yes", 50, false)); //capital D here!!!
                list_bic_mouse.Add(new BuiltInCommand("Show the mousegrid and triple click\r\n" +
                    "in the chosen figure.",
                    bic_type.triple, "triple", null, "No", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Perform a drag and drop operation\r\n" +
                    "by using the mousegrid.",
                    bic_type.drag, "drag", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Perform a left click.",
                    bic_type.click, "click", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Perform a right click.",
                    bic_type.right_click, "east click", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Perform a double click.",
                    bic_type.double_click, "double click", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Perform a triple click.",
                    bic_type.triple_click, "triple click", null, "No", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Hold the left button.",
                    bic_type.hold, "hold", null, "Yes", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Hold the right button.",
                    bic_type.hold_right, "hold east", null, "Yes", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Release mouse buttons if pressed.",
                    bic_type.release_buttons, "release", null, "No", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor up by a chosen number\r\n" +
                    "of pixels (max value: 100).",
                    bic_type.move_up, s_mouse_moves[0] + " pixels_nr", null, "No", 1, true));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor down by a chosen number\r\n" +
                    "of pixels (max value: 100).",
                    bic_type.move_down, s_mouse_moves[1] + " pixels_nr", null, "No", 1, true));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor left by a chosen number\r\n" +
                    "of pixels (max value: 100).",
                    bic_type.move_left, s_mouse_moves[2] + " pixels_nr", null, "No", 1, true));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor right by a chosen number\r\n" +
                    "of pixels (max value: 100).",
                    bic_type.move_right, s_mouse_moves[3] + " pixels_nr", null, "No", 1, true));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor to the middle of " +
                    "the top screen edge.",
                    bic_type.move_top_edge, move_edges_str[0], null, "No", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor to the middle of " +
                    "the bottom screen edge.",
                    bic_type.move_bottom_edge, move_edges_str[1], null, "No", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor to the middle of " +
                    "the left screen edge.",
                    bic_type.move_left_edge, move_edges_str[2], null, "No", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor to the middle of " +
                    "the right screen edge.",
                    bic_type.move_right_edge, move_edges_str[3], null, "No", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Move cursor to the middle of the screen.",
                    bic_type.move_screen_center, "screen center", null, "No", 1, false));
                list_bic_mouse.Add(new BuiltInCommand("Scroll up the mouse wheel.",
                    bic_type.scroll_up, "scroll up", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Scroll down the mouse wheel.",
                    bic_type.scroll_down, "scroll down", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Scroll left the mouse wheel.",
                    bic_type.scroll_left, "scroll left", null, "Yes", 50, false));
                list_bic_mouse.Add(new BuiltInCommand("Scroll right the mouse wheel.",
                    bic_type.scroll_right, "scroll right", null, "Yes", 50, false));

                list_bic_keys_pressing = new List<BuiltInCommand>();

                list_bic_keys_pressing.Add(new BuiltInCommand("Windows key",
                    bic_type.key_combination, "start", "windows", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Windows + tab (say up/down/left/right\r\n" +
                    "to select a window and enter to switch to it)",
                    bic_type.key_combination, "switch application", "windows tab", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Windows + D",
                    bic_type.key_combination, "show desktop", "windows delta", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Windows + E",
                    bic_type.key_combination, "open computer", "windows echo", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + shift + escape",
                    bic_type.key_combination, "open task manager", "control shift escape", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Windows + I",
                    bic_type.key_combination, "open settings", "windows india", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Windows + X",
                    bic_type.key_combination, "open power menu", "windows xray", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("F2",
                    bic_type.key_combination, "rename", "function 2", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + F",
                    bic_type.key_combination, "find", "control foxtrot", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + H (in Word, Notepad, etc.)",
                    bic_type.key_combination, "replace", "control hotel", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Alt + enter (open properties of selected file(s))",
                    bic_type.key_combination, "properties", "alt enter", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Shift + F10 (simulate right-click " +
                    "on selected item(s))",
                    bic_type.key_combination, "menu", "shift function 10", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Alt + print screen (create screenshot\r\n" +
                    "for the current program)",
                    bic_type.key_combination, "capture that", "alt print screen", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + A",
                    bic_type.key_combination, "select all", "control alfa", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + C",
                    bic_type.key_combination, "copy", "control charlie", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + X",
                    bic_type.key_combination, "cut", "control xray", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + V",
                    bic_type.key_combination, "paste", "control victor", "No", 20, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + Z",
                    bic_type.key_combination, "undo", "control zulu", "No", 20, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + Y",
                    bic_type.key_combination, "redo", "control yankee", "No", 20, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + N",
                    bic_type.key_combination, "new", "control november", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + O",
                    bic_type.key_combination, "open", "control oscar", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + S",
                    bic_type.key_combination, "save", "control sierra", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + P",
                    bic_type.key_combination, "print", "control papa", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + L (in web browser)",
                    bic_type.key_combination, "web address", "control lima", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + T (in web browser)",
                    bic_type.key_combination, "new tab", "control tango", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + W (in web browser)",
                    bic_type.key_combination, "close tab", "control whiskey", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + Shift + T (in web browser reopen\r\n" +
                    "previously closed tabs)",
                    bic_type.key_combination, "restore tab", "control shift tango", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + Shift + Tab (in web browser)",
                    bic_type.key_combination, "previous tab", "control shift tab", "No", 10, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + Tab (in web browser)",
                    bic_type.key_combination, "next tab", "control tab", "No", 10, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 1 (in web browser)",
                    bic_type.key_combination, "first tab", "control 1", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 2 (in web browser)",
                    bic_type.key_combination, "second tab", "control 2", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 3 (in web browser)",
                    bic_type.key_combination, "third tab", "control 3", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 4 (in web browser)",
                    bic_type.key_combination, "fourth tab", "control 4", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 5 (in web browser)",
                    bic_type.key_combination, "fifth tab", "control 5", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 6 (in web browser)",
                    bic_type.key_combination, "sixth tab", "control 6", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 7 (in web browser)",
                    bic_type.key_combination, "seventh tab", "control 7", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 8 (in web browser)",
                    bic_type.key_combination, "eight tab", "control 8", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl + 9 (in web browser)",
                    bic_type.key_combination, "last tab", "control 9", "No", 1, false));
                list_bic_keys_pressing.Add(new BuiltInCommand("Browser back key",
                    bic_type.key_pressing, "back", VirtualKeyCode.BROWSER_BACK, "No", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Browser forward key",
                    bic_type.key_pressing, "next", VirtualKeyCode.BROWSER_FORWARD, "No", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Browser search key",
                    bic_type.key_pressing, "search", VirtualKeyCode.BROWSER_SEARCH, "No", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Browser refresh key",
                    bic_type.key_pressing, "refresh", VirtualKeyCode.BROWSER_REFRESH, "No", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Volume up key",
                    bic_type.key_pressing, "volume up", VirtualKeyCode.VOLUME_UP, "No", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Volume down key",
                    bic_type.key_pressing, "volume down", VirtualKeyCode.VOLUME_DOWN, "No", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Volume mute key",
                    bic_type.key_pressing, "volume mute", VirtualKeyCode.VOLUME_MUTE, "No", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Media play pause key",
                    bic_type.key_pressing, "play pause", VirtualKeyCode.MEDIA_PLAY_PAUSE, "No", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Media stop key",
                    bic_type.key_pressing, "stop media", VirtualKeyCode.MEDIA_STOP, "No", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Media next track key",
                    bic_type.key_pressing, "next track", VirtualKeyCode.MEDIA_NEXT_TRACK, "No", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Media previous track key",
                    bic_type.key_pressing, "previous track", VirtualKeyCode.MEDIA_PREV_TRACK, "No", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Ctrl",
                    bic_type.key_pressing, "control", VirtualKeyCode.CONTROL, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Shift",
                    bic_type.key_pressing, "shift", VirtualKeyCode.SHIFT, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Left alt",
                    bic_type.key_pressing, "alt", VirtualKeyCode.LMENU, "Yes", 20)); //LMENU here!!!
                list_bic_keys_pressing.Add(new BuiltInCommand("Right alt",
                    bic_type.key_pressing, "right alt", VirtualKeyCode.RMENU, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Windows key",
                    bic_type.key_pressing, "windows", VirtualKeyCode.LWIN, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Space",
                    bic_type.key_pressing, "space", VirtualKeyCode.SPACE, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Escape (can cancel actions like file\r\n" +
                    "saving, opening, printing, etc.)",
                    bic_type.key_pressing, "escape", VirtualKeyCode.ESCAPE, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Tab",
                    bic_type.key_pressing, "tab", VirtualKeyCode.TAB, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Caps lock",
                    bic_type.key_pressing, "caps lock", VirtualKeyCode.CAPITAL, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Backspace",
                    bic_type.key_pressing, "backspace", VirtualKeyCode.BACK, "Yes", 50));
                list_bic_keys_pressing.Add(new BuiltInCommand("Enter",
                    bic_type.key_pressing, "enter", VirtualKeyCode.RETURN, "Yes", 50));
                list_bic_keys_pressing.Add(new BuiltInCommand("Insert",
                    bic_type.key_pressing, "insert", VirtualKeyCode.INSERT, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Delete",
                    bic_type.key_pressing, "delete", VirtualKeyCode.DELETE, "Yes", 50));
                list_bic_keys_pressing.Add(new BuiltInCommand("Home",
                    bic_type.key_pressing, "home", VirtualKeyCode.HOME, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("End",
                    bic_type.key_pressing, "end", VirtualKeyCode.END, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Page up",
                    bic_type.key_pressing, "page up", VirtualKeyCode.PRIOR, "Yes", 50));
                list_bic_keys_pressing.Add(new BuiltInCommand("Page down",
                    bic_type.key_pressing, "page down", VirtualKeyCode.NEXT, "Yes", 50));
                list_bic_keys_pressing.Add(new BuiltInCommand("Print screen",
                    bic_type.key_pressing, "print screen", VirtualKeyCode.SNAPSHOT, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("Up",
                    bic_type.key_pressing, "up", VirtualKeyCode.UP, "Yes", 50));
                list_bic_keys_pressing.Add(new BuiltInCommand("Down",
                    bic_type.key_pressing, "down", VirtualKeyCode.DOWN, "Yes", 50));
                list_bic_keys_pressing.Add(new BuiltInCommand("Left",
                    bic_type.key_pressing, "left", VirtualKeyCode.LEFT, "Yes", 90));
                list_bic_keys_pressing.Add(new BuiltInCommand("Right",
                    bic_type.key_pressing, "right", VirtualKeyCode.RIGHT, "Yes", 90));
                list_bic_keys_pressing.Add(new BuiltInCommand("1",
                    bic_type.key_pressing, "1", VirtualKeyCode.VK_1, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("2",
                    bic_type.key_pressing, "2", VirtualKeyCode.VK_2, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("3",
                    bic_type.key_pressing, "3", VirtualKeyCode.VK_3, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("4",
                    bic_type.key_pressing, "4", VirtualKeyCode.VK_4, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("5",
                    bic_type.key_pressing, "5", VirtualKeyCode.VK_5, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("6",
                    bic_type.key_pressing, "6", VirtualKeyCode.VK_6, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("7",
                    bic_type.key_pressing, "7", VirtualKeyCode.VK_7, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("8",
                    bic_type.key_pressing, "8", VirtualKeyCode.VK_8, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("9",
                    bic_type.key_pressing, "9", VirtualKeyCode.VK_9, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("0",
                    bic_type.key_pressing, "0", VirtualKeyCode.VK_0, "Yes", 20));
                //needed for people with French r accent:
                list_bic_keys_pressing.Add(new BuiltInCommand("0",
                    bic_type.key_pressing, "null", VirtualKeyCode.VK_0, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("A",
                    bic_type.key_pressing, "alfa", VirtualKeyCode.VK_A, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("B",
                    bic_type.key_pressing, "bravo", VirtualKeyCode.VK_B, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("C",
                    bic_type.key_pressing, "charlie", VirtualKeyCode.VK_C, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("D",
                    bic_type.key_pressing, "delta", VirtualKeyCode.VK_D, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("E",
                    bic_type.key_pressing, "echo", VirtualKeyCode.VK_E, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("F",
                    bic_type.key_pressing, "foxtrot", VirtualKeyCode.VK_F, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("G",
                    bic_type.key_pressing, "golf", VirtualKeyCode.VK_G, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("H",
                    bic_type.key_pressing, "hotel", VirtualKeyCode.VK_H, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("I",
                    bic_type.key_pressing, "india", VirtualKeyCode.VK_I, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("J",
                    bic_type.key_pressing, "juliett", VirtualKeyCode.VK_J, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("K",
                    bic_type.key_pressing, "kilo", VirtualKeyCode.VK_K, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("L",
                    bic_type.key_pressing, "lima", VirtualKeyCode.VK_L, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("M",
                    bic_type.key_pressing, "mike", VirtualKeyCode.VK_M, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("N",
                    bic_type.key_pressing, "november", VirtualKeyCode.VK_N, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("O",
                    bic_type.key_pressing, "oscar", VirtualKeyCode.VK_O, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("P",
                    bic_type.key_pressing, "papa", VirtualKeyCode.VK_P, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Q",
                    bic_type.key_pressing, "quebec", VirtualKeyCode.VK_Q, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("R",
                    bic_type.key_pressing, "romeo", VirtualKeyCode.VK_R, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("S",
                    bic_type.key_pressing, "sierra", VirtualKeyCode.VK_S, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("T",
                    bic_type.key_pressing, "tango", VirtualKeyCode.VK_T, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("U",
                    bic_type.key_pressing, "uniform", VirtualKeyCode.VK_U, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("V",
                    bic_type.key_pressing, "victor", VirtualKeyCode.VK_V, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("W",
                    bic_type.key_pressing, "whiskey", VirtualKeyCode.VK_W, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("X",
                    bic_type.key_pressing, "xray", VirtualKeyCode.VK_X, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Y",
                    bic_type.key_pressing, "yankee", VirtualKeyCode.VK_Y, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("Z",
                    bic_type.key_pressing, "zulu", VirtualKeyCode.VK_Z, "Yes", 20));
                list_bic_keys_pressing.Add(new BuiltInCommand("F1",
                    bic_type.key_pressing, "function 1", VirtualKeyCode.F1, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F2",
                    bic_type.key_pressing, "function 2", VirtualKeyCode.F2, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F3",
                    bic_type.key_pressing, "function 3", VirtualKeyCode.F3, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F4",
                    bic_type.key_pressing, "function 4", VirtualKeyCode.F4, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F5",
                    bic_type.key_pressing, "function 5", VirtualKeyCode.F5, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F6",
                    bic_type.key_pressing, "function 6", VirtualKeyCode.F6, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F7",
                    bic_type.key_pressing, "function 7", VirtualKeyCode.F7, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F8",
                    bic_type.key_pressing, "function 8", VirtualKeyCode.F8, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F9",
                    bic_type.key_pressing, "function 9", VirtualKeyCode.F9, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F10",
                    bic_type.key_pressing, "function 10", VirtualKeyCode.F10, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F11",
                    bic_type.key_pressing, "function 11", VirtualKeyCode.F11, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("F12",
                    bic_type.key_pressing, "function 12", VirtualKeyCode.F12, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand(",",
                    bic_type.key_pressing, "comma", VirtualKeyCode.OEM_COMMA, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand(".",
                    bic_type.key_pressing, "dot", VirtualKeyCode.OEM_PERIOD, "Yes", 3));
                list_bic_keys_pressing.Add(new BuiltInCommand(".",
                    bic_type.key_pressing, "period", VirtualKeyCode.OEM_PERIOD, "Yes", 3));
                list_bic_keys_pressing.Add(new BuiltInCommand("/",
                    bic_type.key_pressing, "slash", VirtualKeyCode.DIVIDE, "Yes", 2));
                list_bic_keys_pressing.Add(new BuiltInCommand("-",
                    bic_type.key_pressing, "dash", VirtualKeyCode.OEM_MINUS, "Yes", 2));
                list_bic_keys_pressing.Add(new BuiltInCommand("-",
                    bic_type.key_pressing, "minus", VirtualKeyCode.OEM_MINUS, "Yes", 2));
                list_bic_keys_pressing.Add(new BuiltInCommand("-",
                    bic_type.key_pressing, "hyphen", VirtualKeyCode.OEM_MINUS, "Yes", 2));
                list_bic_keys_pressing.Add(new BuiltInCommand("*",
                    bic_type.key_pressing, "multiply", VirtualKeyCode.MULTIPLY, "Yes", 1));
                list_bic_keys_pressing.Add(new BuiltInCommand("*",
                    bic_type.key_pressing, "asterisk", VirtualKeyCode.MULTIPLY, "Yes", 1));

                list_bic_char_inserting = new List<BuiltInCommand>();

                list_bic_char_inserting.Add(new BuiltInCommand(";",
                    bic_type.character_ins, "semicolon", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("'",
                    bic_type.character_ins, "quote", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("'",
                    bic_type.character_ins, "apostrophe", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("\\",
                    bic_type.character_ins, "backslash", 2));
                list_bic_char_inserting.Add(new BuiltInCommand("[",
                    bic_type.character_ins, "open bracket", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("]",
                    bic_type.character_ins, "close bracket", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("+",
                    bic_type.character_ins, "plus", 2));
                list_bic_char_inserting.Add(new BuiltInCommand("=",
                    bic_type.character_ins, "equal", 3));
                list_bic_char_inserting.Add(new BuiltInCommand(":",
                    bic_type.character_ins, "colon", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("\"",
                    bic_type.character_ins, "double quote", 1));
                list_bic_char_inserting.Add(new BuiltInCommand(">",
                    bic_type.character_ins, "greater than", 3));
                list_bic_char_inserting.Add(new BuiltInCommand("<",
                    bic_type.character_ins, "less than", 3));
                list_bic_char_inserting.Add(new BuiltInCommand("@",
                    bic_type.character_ins, "at", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("!",
                    bic_type.character_ins, "exclamation", 3));
                list_bic_char_inserting.Add(new BuiltInCommand("?",
                    bic_type.character_ins, "question", 3));
                list_bic_char_inserting.Add(new BuiltInCommand("#",
                    bic_type.character_ins, "number", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("#",
                    bic_type.character_ins, "hash", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("#",
                    bic_type.character_ins, "sharp", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("£",
                    bic_type.character_ins, "pound", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("$",
                    bic_type.character_ins, "dollar", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("%",
                    bic_type.character_ins, "percent", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("^",
                    bic_type.character_ins, "caret", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("(",
                    bic_type.character_ins, "open paren", 1));
                list_bic_char_inserting.Add(new BuiltInCommand(")",
                    bic_type.character_ins, "close paren", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("_",
                    bic_type.character_ins, "underscore", 2));
                list_bic_char_inserting.Add(new BuiltInCommand("{",
                    bic_type.character_ins, "open brace", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("}",
                    bic_type.character_ins, "close brace", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("|",
                    bic_type.character_ins, "vertical bar", 2));
                list_bic_char_inserting.Add(new BuiltInCommand("&",
                    bic_type.character_ins, "ampersand", 2));
                list_bic_char_inserting.Add(new BuiltInCommand("//",
                    bic_type.character_ins, "double slash", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("`",
                    bic_type.character_ins, "back quote", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("±",
                    bic_type.character_ins, "plus or minus", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("×",
                    bic_type.character_ins, "multiplication", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("÷",
                    bic_type.character_ins, "division", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("¢",
                    bic_type.character_ins, "cent", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("§",
                    bic_type.character_ins, "section", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("©",
                    bic_type.character_ins, "copyright", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("®",
                    bic_type.character_ins, "registered", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("°",
                    bic_type.character_ins, "degree", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("¶",
                    bic_type.character_ins, "paragraph", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("ƒ",
                    bic_type.character_ins, "function", 1));
                list_bic_char_inserting.Add(new BuiltInCommand("µ",
                    bic_type.character_ins, "micro", 1));

                list_bic_dictation_always = new List<BuiltInCommand>();

                list_bic_dictation_always.Add(new BuiltInCommand("Turn off speech recognition",
                    bic_type.turn_off, turn_off, 1));
                list_bic_dictation_always.Add(new BuiltInCommand("Switch to command mode",
                    bic_type.switch_to_command, switch_to_command_mode, 1));

                list_bic_dictation_better = new List<BuiltInCommand>();

                list_bic_dictation_better.Add(new BuiltInCommand("Change Windows Dictation Tool mode\r\n" +
                    "to listening by restarting it.",
                    bic_type.start_better_dictation_listening, start_better_dictation_listening, 1));
                list_bic_dictation_better.Add(new BuiltInCommand("Toggle Windows Dictation Tool",
                    bic_type.toggle_better_dictation, toggle_better_dictation_str, 1));

                LVbic_off.ItemsSource = list_bic_off;
                LVbic_general.ItemsSource = list_bic_general;
                LVbic_mouse.ItemsSource = list_bic_mouse;
                LVbic_keys_pressing.ItemsSource = list_bic_keys_pressing;
                LVbic_char_inserting.ItemsSource = list_bic_char_inserting;
                LVbic_dict_always.ItemsSource = list_bic_dictation_always;
                LVbic_dict_better.ItemsSource = list_bic_dictation_better;

                cv_bic_off =
                    (CollectionView)CollectionViewSource.GetDefaultView(LVbic_off.ItemsSource);
                cv_bic_general =
                    (CollectionView)CollectionViewSource.GetDefaultView(LVbic_general.ItemsSource);
                cv_bic_mouse =
                    (CollectionView)CollectionViewSource.GetDefaultView(LVbic_mouse.ItemsSource);
                cv_bic_pressing =
                    (CollectionView)CollectionViewSource.GetDefaultView(LVbic_keys_pressing.ItemsSource);
                cv_bic_inserting =
                    (CollectionView)CollectionViewSource.GetDefaultView(LVbic_char_inserting.ItemsSource);
                cv_bic_dict_always =
                    (CollectionView)CollectionViewSource.GetDefaultView(LVbic_dict_always.ItemsSource);
                cv_bic_dict_better =
                    (CollectionView)CollectionViewSource.GetDefaultView(LVbic_dict_better.ItemsSource);

                create_bic_general_and_mouse_list();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void create_bic_general_and_mouse_list()
        {
            try
            {
                list_bic_general_and_mouse = new List<BuiltInCommand>();
                string name;

                foreach (BuiltInCommand bic in list_bic_general)
                {
                    if (bic.type == bic_type.switch_to_app)
                    {
                        name = switch_to_app_str;
                    }
                    else if (bic.type == bic_type.open_app)
                    {
                        name = open_app_str;
                    }
                    else
                    {
                        name = bic.name;
                    }

                    list_bic_general_and_mouse.Add(new BuiltInCommand()
                    {
                        description = bic.description,
                        enabled = bic.enabled,
                        max_executions = bic.max_executions,
                        key_combination = bic.key_combination,
                        keys = bic.keys,
                        name = name,
                        name_firstupper = bic.name_firstupper,
                        type = bic.type,
                        use_contains = bic.use_contains,
                        vkc = bic.vkc
                    });
                }

                foreach (BuiltInCommand bic in list_bic_mouse)
                {
                    list_bic_general_and_mouse.Add(new BuiltInCommand()
                    {
                        description = bic.description,
                        enabled = bic.enabled,
                        max_executions = bic.max_executions,
                        key_combination = bic.key_combination,
                        keys = bic.keys,
                        name = bic.name,
                        name_firstupper = bic.name_firstupper,
                        type = bic.type,
                        use_contains = bic.use_contains,
                        vkc = bic.vkc
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void save_bic_toggling_data()
        {
            try
            {
                XmlDocument xml_doc = new XmlDocument();

                XmlNode root_node = xml_doc.CreateElement("built-in_commands");

                XmlAttribute attribute = xml_doc.CreateAttribute("version");
                attribute.Value = "1";
                root_node.Attributes.Append(attribute);

                xml_doc.AppendChild(root_node);

                {
                    XmlNode list_node = xml_doc.CreateElement("off");

                    root_node.AppendChild(list_node);

                    foreach (BuiltInCommand bic in list_bic_off)
                    {
                        XmlNode bic_node = xml_doc.CreateElement("command");

                        XmlNode name_node = xml_doc.CreateElement("name");
                        name_node.InnerText = bic.name_firstupper;

                        XmlNode enabled_node = xml_doc.CreateElement("enabled");
                        enabled_node.InnerText = bic.enabled.ToString();

                        bic_node.AppendChild(name_node);
                        bic_node.AppendChild(enabled_node);

                        list_node.AppendChild(bic_node);
                    }
                }
                {
                    XmlNode list_node = xml_doc.CreateElement("general");

                    root_node.AppendChild(list_node);

                    foreach (BuiltInCommand bic in list_bic_general)
                    {
                        XmlNode bic_node = xml_doc.CreateElement("command");

                        XmlNode name_node = xml_doc.CreateElement("name");
                        name_node.InnerText = bic.name_firstupper;

                        XmlNode enabled_node = xml_doc.CreateElement("enabled");
                        enabled_node.InnerText = bic.enabled.ToString();

                        bic_node.AppendChild(name_node);
                        bic_node.AppendChild(enabled_node);

                        list_node.AppendChild(bic_node);
                    }
                }
                {
                    XmlNode list_node = xml_doc.CreateElement("mouse");

                    root_node.AppendChild(list_node);

                    foreach (BuiltInCommand bic in list_bic_mouse)
                    {
                        XmlNode bic_node = xml_doc.CreateElement("command");

                        XmlNode name_node = xml_doc.CreateElement("name");
                        name_node.InnerText = bic.name_firstupper;

                        XmlNode enabled_node = xml_doc.CreateElement("enabled");
                        enabled_node.InnerText = bic.enabled.ToString();

                        bic_node.AppendChild(name_node);
                        bic_node.AppendChild(enabled_node);

                        list_node.AppendChild(bic_node);
                    }
                }
                {
                    XmlNode list_node = xml_doc.CreateElement("keys_pressing");

                    root_node.AppendChild(list_node);

                    foreach (BuiltInCommand bic in list_bic_keys_pressing)
                    {
                        XmlNode bic_node = xml_doc.CreateElement("command");

                        XmlNode name_node = xml_doc.CreateElement("name");
                        name_node.InnerText = bic.name_firstupper;

                        XmlNode enabled_node = xml_doc.CreateElement("enabled");
                        enabled_node.InnerText = bic.enabled.ToString();

                        bic_node.AppendChild(name_node);
                        bic_node.AppendChild(enabled_node);

                        list_node.AppendChild(bic_node);
                    }
                }
                {
                    XmlNode list_node = xml_doc.CreateElement("character_inserting");

                    root_node.AppendChild(list_node);

                    foreach (BuiltInCommand bic in list_bic_char_inserting)
                    {
                        XmlNode bic_node = xml_doc.CreateElement("command");

                        XmlNode name_node = xml_doc.CreateElement("name");
                        name_node.InnerText = bic.name_firstupper;

                        XmlNode enabled_node = xml_doc.CreateElement("enabled");
                        enabled_node.InnerText = bic.enabled.ToString();

                        bic_node.AppendChild(name_node);
                        bic_node.AppendChild(enabled_node);

                        list_node.AppendChild(bic_node);
                    }
                }
                {
                    XmlNode list_node = xml_doc.CreateElement("dictation_always");

                    root_node.AppendChild(list_node);

                    foreach (BuiltInCommand bic in list_bic_dictation_always)
                    {
                        XmlNode bic_node = xml_doc.CreateElement("command");

                        XmlNode name_node = xml_doc.CreateElement("name");
                        name_node.InnerText = bic.name_firstupper;

                        XmlNode enabled_node = xml_doc.CreateElement("enabled");
                        enabled_node.InnerText = bic.enabled.ToString();

                        bic_node.AppendChild(name_node);
                        bic_node.AppendChild(enabled_node);

                        list_node.AppendChild(bic_node);
                    }
                }
                {
                    XmlNode list_node = xml_doc.CreateElement("dictation_better");

                    root_node.AppendChild(list_node);

                    foreach (BuiltInCommand bic in list_bic_dictation_better)
                    {
                        XmlNode bic_node = xml_doc.CreateElement("command");

                        XmlNode name_node = xml_doc.CreateElement("name");
                        name_node.InnerText = bic.name_firstupper;

                        XmlNode enabled_node = xml_doc.CreateElement("enabled");
                        enabled_node.InnerText = bic.enabled.ToString();

                        bic_node.AppendChild(name_node);
                        bic_node.AppendChild(enabled_node);

                        list_node.AppendChild(bic_node);
                    }
                }

                if (Directory.Exists(Middle_Man.saving_folder_path) == false)
                {
                    Directory.CreateDirectory(Middle_Man.saving_folder_path);
                }

                xml_doc.Save(Path.Combine(Middle_Man.saving_folder_path, Middle_Man.bic_filename));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void load_bic_toggling_data()
        {
            try
            {
                if (File.Exists(Path.Combine(Middle_Man.saving_folder_path, Middle_Man.bic_filename)))
                {
                    updating_bic_toggling_data_on = false;

                    XmlDocument xml_doc = new XmlDocument();
                    xml_doc.Load(Path.Combine(Middle_Man.saving_folder_path, Middle_Man.bic_filename));

                    XmlNodeList bic = xml_doc.SelectNodes("//built-in_commands");

                    int version = -1;
                    bool parsing_v = false;

                    if (bic[0].Attributes["version"] != null)
                        parsing_v = int.TryParse(bic[0].Attributes["version"].Value, out version);

                    if (parsing_v && version == 1)
                    {
                        XmlNodeList commands = xml_doc.SelectNodes("//built-in_commands/off")[0].ChildNodes;

                        foreach (XmlNode command in commands)
                        {
                            XmlNodeList nodes = command.ChildNodes;

                            string name = "";
                            bool enabled = true;

                            int i = 0;

                            foreach (XmlNode node in nodes)
                            {
                                if (node.Name == "name")
                                {
                                    name = node.InnerText;
                                    i++;
                                }
                                else if (node.Name == "enabled")
                                {
                                    bool parsing = bool.TryParse(node.InnerText, out enabled);

                                    if (parsing)
                                        i++;
                                }
                            }

                            if (i > 1)
                            {
                                toggle_bic_off(enabled, name);
                            }
                        }

                        commands = xml_doc.SelectNodes("//built-in_commands/general")[0].ChildNodes;

                        foreach (XmlNode command in commands)
                        {
                            XmlNodeList nodes = command.ChildNodes;

                            string name = "";
                            bool enabled = true;

                            int i = 0;

                            foreach (XmlNode node in nodes)
                            {
                                if (node.Name == "name")
                                {
                                    name = node.InnerText;
                                    i++;
                                }
                                else if (node.Name == "enabled")
                                {
                                    bool parsing = bool.TryParse(node.InnerText, out enabled);

                                    if (parsing)
                                        i++;
                                }
                            }

                            if (i > 1)
                            {
                                toggle_bic_general(enabled, name);
                            }
                        }

                        commands = xml_doc.SelectNodes("//built-in_commands/mouse")[0].ChildNodes;

                        foreach (XmlNode command in commands)
                        {
                            XmlNodeList nodes = command.ChildNodes;

                            string name = "";
                            bool enabled = true;

                            int i = 0;

                            foreach (XmlNode node in nodes)
                            {
                                if (node.Name == "name")
                                {
                                    name = node.InnerText;
                                    i++;
                                }
                                else if (node.Name == "enabled")
                                {
                                    bool parsing = bool.TryParse(node.InnerText, out enabled);

                                    if (parsing)
                                        i++;
                                }
                            }

                            if (i > 1)
                            {
                                toggle_bic_mouse(enabled, name);
                            }
                        }

                        commands = xml_doc.SelectNodes("//built-in_commands/keys_pressing")[0].ChildNodes;

                        foreach (XmlNode command in commands)
                        {
                            XmlNodeList nodes = command.ChildNodes;

                            string name = "";
                            bool enabled = true;

                            int i = 0;

                            foreach (XmlNode node in nodes)
                            {
                                if (node.Name == "name")
                                {
                                    name = node.InnerText;
                                    i++;
                                }
                                else if (node.Name == "enabled")
                                {
                                    bool parsing = bool.TryParse(node.InnerText, out enabled);

                                    if (parsing)
                                        i++;
                                }
                            }

                            if (i > 1)
                            {
                                toggle_bic_pressing(enabled, name);
                            }
                        }

                        commands = xml_doc.SelectNodes("//built-in_commands/character_inserting")[0].ChildNodes;

                        foreach (XmlNode command in commands)
                        {
                            XmlNodeList nodes = command.ChildNodes;

                            string name = "";
                            bool enabled = true;

                            int i = 0;

                            foreach (XmlNode node in nodes)
                            {
                                if (node.Name == "name")
                                {
                                    name = node.InnerText;
                                    i++;
                                }
                                else if (node.Name == "enabled")
                                {
                                    bool parsing = bool.TryParse(node.InnerText, out enabled);

                                    if (parsing)
                                        i++;
                                }
                            }

                            if (i > 1)
                            {
                                toggle_bic_inserting(enabled, name);
                            }
                        }

                        commands = xml_doc.SelectNodes("//built-in_commands/dictation_always")[0].ChildNodes;

                        foreach (XmlNode command in commands)
                        {
                            XmlNodeList nodes = command.ChildNodes;

                            string name = "";
                            bool enabled = true;

                            int i = 0;

                            foreach (XmlNode node in nodes)
                            {
                                if (node.Name == "name")
                                {
                                    name = node.InnerText;
                                    i++;
                                }
                                else if (node.Name == "enabled")
                                {
                                    bool parsing = bool.TryParse(node.InnerText, out enabled);

                                    if (parsing)
                                        i++;
                                }
                            }

                            if (i > 1)
                            {
                                toggle_bic_always(enabled, name);
                            }
                        }

                        commands = xml_doc.SelectNodes("//built-in_commands/dictation_better")[0].ChildNodes;

                        foreach (XmlNode command in commands)
                        {
                            XmlNodeList nodes = command.ChildNodes;

                            string name = "";
                            bool enabled = true;

                            int i = 0;

                            foreach (XmlNode node in nodes)
                            {
                                if (node.Name == "name")
                                {
                                    name = node.InnerText;
                                    i++;
                                }
                                else if (node.Name == "enabled")
                                {
                                    bool parsing = bool.TryParse(node.InnerText, out enabled);

                                    if (parsing)
                                        i++;
                                }
                            }

                            if (i > 1)
                            {
                                toggle_bic_better(enabled, name);
                            }
                        }

                        create_bic_general_and_mouse_list();
                    }

                    updating_bic_toggling_data_on = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        bool is_bic_in_general_and_mouse_enabled(bic_type bt)
        {
            try
            {
                foreach (BuiltInCommand bic in list_bic_general_and_mouse)
                {
                    if (bic.type == bt)
                    {
                        if (bic.enabled)
                            return true;
                        else
                            return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC006", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return false;
        }

        bool is_bic_in_dictation_enabled(bic_type bt)
        {
            try
            {
                foreach (BuiltInCommand bic in list_bic_dictation_always)
                {
                    if (bic.type == bt)
                    {
                        if (bic.enabled)
                            return true;
                        else
                            return false;
                    }
                }

                foreach (BuiltInCommand bic in list_bic_dictation_better)
                {
                    if (bic.type == bt)
                    {
                        if (bic.enabled)
                            return true;
                        else
                            return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC006", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return false;
        }

        bool are_all_bic_off_disabled()
        {
            try
            {
                foreach (BuiltInCommand bic in list_bic_off)
                {
                    if (bic.enabled)
                        return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC056", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return true;
        }

        bool are_all_bic_general_and_mouse_disabled()
        {
            try
            {
                foreach (BuiltInCommand bic in list_bic_general_and_mouse)
                {
                    if (bic.enabled)
                        return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC057", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return true;
        }

        bool are_all_bic_dictation_disabled()
        {
            try
            {
                foreach (BuiltInCommand bic in list_bic_dictation_always)
                {
                    if (bic.enabled)
                        return false;
                }

                if (better_dictation)
                {
                    foreach (BuiltInCommand bic in list_bic_dictation_better)
                    {
                        if (bic.enabled)
                            return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC058", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return true;
        }

        void toggle_bic_off(bool enabled, string name_firstupper)
        {
            try
            {
                for (int i = 0; i < list_bic_off.Count; i++)
                {
                    if (list_bic_off[i].name_firstupper == name_firstupper)
                    {
                        list_bic_off[i].enabled = enabled;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC007", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void toggle_bic_general_and_mouse(bool enabled, string name_firstupper)
        {
            try
            {
                for (int i = 0; i < list_bic_general_and_mouse.Count; i++)
                {
                    if (list_bic_general_and_mouse[i].name_firstupper == name_firstupper)
                    {
                        list_bic_general_and_mouse[i].enabled = enabled;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC008", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void toggle_bic_general(bool enabled, string name_firstupper)
        {
            try
            {
                for (int i = 0; i < list_bic_general.Count; i++)
                {
                    if (list_bic_general[i].name_firstupper == name_firstupper)
                    {
                        list_bic_general[i].enabled = enabled;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC009", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void toggle_bic_mouse(bool enabled, string name_firstupper)
        {
            try
            {
                for (int i = 0; i < list_bic_mouse.Count; i++)
                {
                    if (list_bic_mouse[i].name_firstupper == name_firstupper)
                    {
                        list_bic_mouse[i].enabled = enabled;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC010", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void toggle_bic_pressing(bool enabled, string name_firstupper)
        {
            try
            {
                for (int i = 0; i < list_bic_keys_pressing.Count; i++)
                {
                    if (list_bic_keys_pressing[i].name_firstupper == name_firstupper)
                    {
                        list_bic_keys_pressing[i].enabled = enabled;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC011", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void toggle_bic_inserting(bool enabled, string name_firstupper)
        {
            try
            {
                for (int i = 0; i < list_bic_char_inserting.Count; i++)
                {
                    if (list_bic_char_inserting[i].name_firstupper == name_firstupper)
                    {
                        list_bic_char_inserting[i].enabled = enabled;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC012", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void toggle_bic_always(bool enabled, string name_firstupper)
        {
            try
            {
                for (int i = 0; i < list_bic_dictation_always.Count; i++)
                {
                    if (list_bic_dictation_always[i].name_firstupper == name_firstupper)
                    {
                        list_bic_dictation_always[i].enabled = enabled;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC013", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void toggle_bic_better(bool enabled, string name_firstupper)
        {
            try
            {
                for (int i = 0; i < list_bic_dictation_better.Count; i++)
                {
                    if (list_bic_dictation_better[i].name_firstupper == name_firstupper)
                    {
                        list_bic_dictation_better[i].enabled = enabled;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC014", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_off_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                if (LVbic_off.SelectedIndex != -1)
                {
                    Benable_bic_off.IsEnabled = true;
                    Bdisable_bic_off.IsEnabled = true;
                    MIenable_bic_off.IsEnabled = true;
                    MIdisable_bic_off.IsEnabled = true;
                }
                else
                {
                    Benable_bic_off.IsEnabled = false;
                    Bdisable_bic_off.IsEnabled = false;
                    MIenable_bic_off.IsEnabled = false;
                    MIdisable_bic_off.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC015", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_general_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                if (LVbic_general.SelectedIndex != -1)
                {
                    Benable_bic_general.IsEnabled = true;
                    Bdisable_bic_general.IsEnabled = true;
                    MIenable_bic_general.IsEnabled = true;
                    MIdisable_bic_general.IsEnabled = true;
                }
                else
                {
                    Benable_bic_general.IsEnabled = false;
                    Bdisable_bic_general.IsEnabled = false;
                    MIenable_bic_general.IsEnabled = false;
                    MIdisable_bic_general.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC016", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_mouse_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                if (LVbic_mouse.SelectedIndex != -1)
                {
                    Benable_bic_mouse.IsEnabled = true;
                    Bdisable_bic_mouse.IsEnabled = true;
                    MIenable_bic_mouse.IsEnabled = true;
                    MIdisable_bic_mouse.IsEnabled = true;
                }
                else
                {
                    Benable_bic_mouse.IsEnabled = false;
                    Bdisable_bic_mouse.IsEnabled = false;
                    MIenable_bic_mouse.IsEnabled = false;
                    MIdisable_bic_mouse.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC017", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_keys_pressing_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                if (LVbic_keys_pressing.SelectedIndex != -1)
                {
                    Benable_bic_pressing.IsEnabled = true;
                    Bdisable_bic_pressing.IsEnabled = true;
                    MIenable_bic_pressing.IsEnabled = true;
                    MIdisable_bic_pressing.IsEnabled = true;
                }
                else
                {
                    Benable_bic_pressing.IsEnabled = false;
                    Bdisable_bic_pressing.IsEnabled = false;
                    MIenable_bic_pressing.IsEnabled = false;
                    MIdisable_bic_pressing.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC018", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_char_inserting_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                if (LVbic_char_inserting.SelectedIndex != -1)
                {
                    Benable_bic_inserting.IsEnabled = true;
                    Bdisable_bic_inserting.IsEnabled = true;
                    MIenable_bic_inserting.IsEnabled = true;
                    MIdisable_bic_inserting.IsEnabled = true;
                }
                else
                {
                    Benable_bic_inserting.IsEnabled = false;
                    Bdisable_bic_inserting.IsEnabled = false;
                    MIenable_bic_inserting.IsEnabled = false;
                    MIdisable_bic_inserting.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC019", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_dict_always_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                if (LVbic_dict_always.SelectedIndex != -1)
                {
                    Benable_bic_always.IsEnabled = true;
                    Bdisable_bic_always.IsEnabled = true;
                    MIenable_bic_always.IsEnabled = true;
                    MIdisable_bic_always.IsEnabled = true;
                }
                else
                {
                    Benable_bic_always.IsEnabled = false;
                    Bdisable_bic_always.IsEnabled = false;
                    MIenable_bic_always.IsEnabled = false;
                    MIdisable_bic_always.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC020", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_dict_better_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                if (LVbic_dict_better.SelectedIndex != -1)
                {
                    Benable_bic_better.IsEnabled = true;
                    Bdisable_bic_better.IsEnabled = true;
                    MIenable_bic_better.IsEnabled = true;
                    MIdisable_bic_better.IsEnabled = true;
                }
                else
                {
                    Benable_bic_better.IsEnabled = false;
                    Bdisable_bic_better.IsEnabled = false;
                    MIenable_bic_better.IsEnabled = false;
                    MIdisable_bic_better.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC021", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Benable_bic_off_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_off.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_off.SelectedItems)
                    {
                        toggle_bic_off(true, bic.name_firstupper);
                    }

                    cv_bic_off.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC022", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Benable_bic_general_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_general.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_general.SelectedItems)
                    {
                        toggle_bic_general(true, bic.name_firstupper);
                        toggle_bic_general_and_mouse(true, bic.name_firstupper);
                    }

                    cv_bic_general.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC023", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Benable_bic_mouse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_mouse.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_mouse.SelectedItems)
                    {
                        toggle_bic_mouse(true, bic.name_firstupper);
                        toggle_bic_general_and_mouse(true, bic.name_firstupper);
                    }

                    cv_bic_mouse.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC024", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Benable_bic_pressing_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_keys_pressing.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_keys_pressing.SelectedItems)
                    {
                        toggle_bic_pressing(true, bic.name_firstupper);
                    }

                    cv_bic_pressing.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC025", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Benable_bic_inserting_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_char_inserting.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_char_inserting.SelectedItems)
                    {
                        toggle_bic_inserting(true, bic.name_firstupper);
                    }

                    cv_bic_inserting.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC026", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Benable_bic_always_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_dict_always.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_dict_always.SelectedItems)
                    {
                        toggle_bic_always(true, bic.name_firstupper);
                    }

                    cv_bic_dict_always.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC026", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Benable_bic_better_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_dict_better.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_dict_better.SelectedItems)
                    {
                        toggle_bic_better(true, bic.name_firstupper);
                    }

                    cv_bic_dict_better.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC027", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdisable_bic_off_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_off.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_off.SelectedItems)
                    {
                        toggle_bic_off(false, bic.name_firstupper);
                    }

                    cv_bic_off.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC028", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdisable_bic_general_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_general.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_general.SelectedItems)
                    {
                        toggle_bic_general(false, bic.name_firstupper);
                        toggle_bic_general_and_mouse(false, bic.name_firstupper);
                    }

                    cv_bic_general.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC029", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdisable_bic_mouse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_mouse.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_mouse.SelectedItems)
                    {
                        toggle_bic_mouse(false, bic.name_firstupper);
                        toggle_bic_general_and_mouse(false, bic.name_firstupper);
                    }

                    cv_bic_mouse.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC030", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdisable_bic_pressing_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_keys_pressing.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_keys_pressing.SelectedItems)
                    {
                        toggle_bic_pressing(false, bic.name_firstupper);
                    }

                    cv_bic_pressing.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC031", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdisable_bic_inserting_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_char_inserting.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_char_inserting.SelectedItems)
                    {
                        toggle_bic_inserting(false, bic.name_firstupper);
                    }

                    cv_bic_inserting.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC032", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdisable_bic_always_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_dict_always.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_dict_always.SelectedItems)
                    {
                        toggle_bic_always(false, bic.name_firstupper);
                    }

                    cv_bic_dict_always.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC033", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdisable_bic_better_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVbic_dict_better.SelectedIndex != -1)
                {
                    foreach (BuiltInCommand bic in LVbic_dict_better.SelectedItems)
                    {
                        toggle_bic_better(false, bic.name_firstupper);
                    }

                    cv_bic_dict_better.Refresh();

                    update_bic_grammar();

                    save_bic_toggling_data();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC034", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bselect_all_bic_off_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LVbic_off.SelectAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC035", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bselect_all_bic_general_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LVbic_general.SelectAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC036", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bselect_all_bic_mouse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LVbic_mouse.SelectAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC037", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bselect_all_bic_pressing_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LVbic_keys_pressing.SelectAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC038", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bselect_all_bic_inserting_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LVbic_char_inserting.SelectAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC039", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bselect_all_bic_better_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LVbic_dict_better.SelectAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC040", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bselect_all_bic_always_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LVbic_dict_always.SelectAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC041", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_off_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.OemPlus && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Benable_bic_off_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Bdisable_bic_off_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC042", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_general_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.OemPlus && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Benable_bic_general_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Bdisable_bic_general_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC043", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_mouse_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.OemPlus && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Benable_bic_mouse_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Bdisable_bic_mouse_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC044", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_keys_pressing_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.OemPlus && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Benable_bic_pressing_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Bdisable_bic_pressing_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC045", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_char_inserting_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.OemPlus && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Benable_bic_inserting_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Bdisable_bic_inserting_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC046", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_dict_always_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.OemPlus && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Benable_bic_always_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Bdisable_bic_always_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC047", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_dict_better_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.OemPlus && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Benable_bic_better_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    Bdisable_bic_better_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC048", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_off_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                BuiltInCommand bic = (BuiltInCommand)LVbic_off.SelectedItem;

                if (bic.enabled)
                    toggle_bic_off(false, bic.name_firstupper);
                else
                    toggle_bic_off(true, bic.name_firstupper);

                cv_bic_off.Refresh();

                update_bic_grammar();

                save_bic_toggling_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC049", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_general_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                BuiltInCommand bic = (BuiltInCommand)LVbic_general.SelectedItem;

                if (bic.enabled)
                {
                    toggle_bic_general(false, bic.name_firstupper);
                    toggle_bic_general_and_mouse(false, bic.name_firstupper);
                }
                else
                {
                    toggle_bic_general(true, bic.name_firstupper);
                    toggle_bic_general_and_mouse(true, bic.name_firstupper);
                }

                cv_bic_general.Refresh();

                update_bic_grammar();

                save_bic_toggling_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC050", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_mouse_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                BuiltInCommand bic = (BuiltInCommand)LVbic_mouse.SelectedItem;

                if (bic.enabled)
                {
                    toggle_bic_mouse(false, bic.name_firstupper);
                    toggle_bic_general_and_mouse(false, bic.name_firstupper);
                }
                else
                {
                    toggle_bic_mouse(true, bic.name_firstupper);
                    toggle_bic_general_and_mouse(true, bic.name_firstupper);
                }

                cv_bic_mouse.Refresh();

                update_bic_grammar();

                save_bic_toggling_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC051", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_keys_pressing_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                BuiltInCommand bic = (BuiltInCommand)LVbic_keys_pressing.SelectedItem;

                if (bic.enabled)
                    toggle_bic_pressing(false, bic.name_firstupper);
                else
                    toggle_bic_pressing(true, bic.name_firstupper);

                cv_bic_pressing.Refresh();

                update_bic_grammar();

                save_bic_toggling_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC052", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_char_inserting_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                BuiltInCommand bic = (BuiltInCommand)LVbic_char_inserting.SelectedItem;

                if (bic.enabled)
                    toggle_bic_inserting(false, bic.name_firstupper);
                else
                    toggle_bic_inserting(true, bic.name_firstupper);

                cv_bic_inserting.Refresh();

                update_bic_grammar();

                save_bic_toggling_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC053", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_dict_always_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                BuiltInCommand bic = (BuiltInCommand)LVbic_dict_always.SelectedItem;

                if (bic.enabled)
                    toggle_bic_always(false, bic.name_firstupper);
                else
                    toggle_bic_always(true, bic.name_firstupper);

                cv_bic_dict_always.Refresh();

                update_bic_grammar();

                save_bic_toggling_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC054", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVbic_dict_better_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                BuiltInCommand bic = (BuiltInCommand)LVbic_dict_better.SelectedItem;

                if (bic.enabled)
                    toggle_bic_better(false, bic.name_firstupper);
                else
                    toggle_bic_better(true, bic.name_firstupper);

                cv_bic_dict_better.Refresh();

                update_bic_grammar();

                save_bic_toggling_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error BIC055", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIenable_bic_off_Click(object sender, RoutedEventArgs e)
        {
            Benable_bic_off_Click(null, null);
        }

        private void MIenable_bic_general_Click(object sender, RoutedEventArgs e)
        {
            Benable_bic_general_Click(null, null);
        }

        private void MIenable_bic_mouse_Click(object sender, RoutedEventArgs e)
        {
            Benable_bic_mouse_Click(null, null);
        }

        private void MIenable_bic_pressing_Click(object sender, RoutedEventArgs e)
        {
            Benable_bic_pressing_Click(null, null);
        }

        private void MIenable_bic_inserting_Click(object sender, RoutedEventArgs e)
        {
            Benable_bic_inserting_Click(null, null);
        }

        private void MIenable_bic_always_Click(object sender, RoutedEventArgs e)
        {
            Benable_bic_always_Click(null, null);
        }

        private void MIenable_bic_better_Click(object sender, RoutedEventArgs e)
        {
            Benable_bic_better_Click(null, null);
        }

        private void MIdisable_bic_off_Click(object sender, RoutedEventArgs e)
        {
            Bdisable_bic_off_Click(null, null);
        }

        private void MIdisable_bic_general_Click(object sender, RoutedEventArgs e)
        {
            Bdisable_bic_general_Click(null, null);
        }

        private void MIdisable_bic_mouse_Click(object sender, RoutedEventArgs e)
        {
            Bdisable_bic_mouse_Click(null, null);
        }

        private void MIdisable_bic_pressing_Click(object sender, RoutedEventArgs e)
        {
            Bdisable_bic_pressing_Click(null, null);
        }

        private void MIdisable_bic_inserting_Click(object sender, RoutedEventArgs e)
        {
            Bdisable_bic_inserting_Click(null, null);
        }

        private void MIdisable_bic_always_Click(object sender, RoutedEventArgs e)
        {
            Bdisable_bic_always_Click(null, null);
        }

        private void MIdisable_bic_better_Click(object sender, RoutedEventArgs e)
        {
            Bdisable_bic_better_Click(null, null);
        }

        private void MIselect_all_bic_off_Click(object sender, RoutedEventArgs e)
        {
            Bselect_all_bic_off_Click(null, null);
        }

        private void MIselect_all_bic_general_Click(object sender, RoutedEventArgs e)
        {
            Bselect_all_bic_general_Click(null, null);
        }

        private void MIselect_all_bic_mouse_Click(object sender, RoutedEventArgs e)
        {
            Bselect_all_bic_mouse_Click(null, null);
        }

        private void MIselect_all_bic_pressing_Click(object sender, RoutedEventArgs e)
        {
            Bselect_all_bic_pressing_Click(null, null);
        }

        private void MIselect_all_bic_inserting_Click(object sender, RoutedEventArgs e)
        {
            Bselect_all_bic_inserting_Click(null, null);
        }

        private void MIselect_all_bic_always_Click(object sender, RoutedEventArgs e)
        {
            Bselect_all_bic_always_Click(null, null);
        }

        private void MIselect_all_bic_better_Click(object sender, RoutedEventArgs e)
        {
            Bselect_all_bic_better_Click(null, null);
        }
    }
}