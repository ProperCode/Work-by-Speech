//highest error nr: CC028
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using WindowsInput.Native;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Speech
{
    [Serializable]
    public class CustomCommand
    {
        public string name { get; set; }
        public string description { get; set; }
        public string group { get; set; }
        public short max_executions { get; set; }
        public bool enabled { get; set; }
        public List<CC_Action> actions { get; set; }
    }

    public class Group 
    {
        public string name { get; set; }
    }

    public class Profile
    {
        public string name { get; set; }
        public string program { get; set; }
        public bool enabled { get; set; }

        public List<CustomCommand> custom_commands { get; set; }
    }

    [Serializable]
    public class CC_Action
    {
        public string action { get; set; }
    }

    public class VirtualKey
    {
        public string name { get; set; }
        public VirtualKeyCode vkc { get; set; }
    }

    public class ActionKeyboard
    {
        public short option { get; set; }
        public int time { get; set; }
        public string[] keys { get; set; }
        public string action_text { get; set; }

        public ActionKeyboard()
        {
        }

        public ActionKeyboard(string Action_text)
        {
            action_text = Action_text;

            if (action_text.StartsWith("Press:"))
            {
                //Press: A + B for 75ms
                option = 0;

                string str = action_text.Replace("Press: ", "");

                string[] arr = str.Split(new string[] { " for " },
                    StringSplitOptions.RemoveEmptyEntries);

                time = 75;

                if (arr.Length > 1)
                {
                    if (int.TryParse(arr[1].Replace("ms", ""), out int n))
                    {
                        if (n > 0)
                            time = n;
                    }
                }

                keys = arr[0].Split(new string[] { " + " },
                    StringSplitOptions.RemoveEmptyEntries);
            }
            else if (action_text.StartsWith("Toggle:"))
            {
                option = 1;

                //Toggle: A + B
                Action_text = Action_text.Replace("Toggle: ", "");

                keys = Action_text.Split(new string[] { " + " },
                    StringSplitOptions.RemoveEmptyEntries);
            }
            else if (action_text.StartsWith("Hold down:"))
            {
                option = 2;

                //Hold down: A + B
                Action_text = Action_text.Replace("Hold down: ", "");

                keys = Action_text.Split(new string[] { " + " },
                    StringSplitOptions.RemoveEmptyEntries);
            }
            else if (action_text.StartsWith("Release:"))
            {
                option = 3;

                //Release: A + B
                Action_text = Action_text.Replace("Release: ", "");

                keys = Action_text.Split(new string[] { " + " },
                    StringSplitOptions.RemoveEmptyEntries);
            }
        }
    }

    public class ActionMouse
    {
        public bool left { get; set; }
        public short option { get; set; }
        public int time { get; set; }
        public string[] keys { get; set; }
        public int x { get; set; }
        public int y { get; set; }
        public string action_text { get; set; }

        public ActionMouse()
        {
        }

        public ActionMouse(string Action_text)
        {
            action_text = Action_text;

            if (action_text.StartsWith("Click: LMB")
                || action_text.StartsWith("Click: RMB"))
            {
                //Click: LMB (x, y) while pressing alt + control for 75ms

                option = 0;
                
                string str = action_text.Replace("Click: ", "");

                left = true;

                if (str.StartsWith("RMB"))
                    left = false;

                string[] arr = str.Split(new string[] { " for " }, StringSplitOptions.RemoveEmptyEntries);

                time = 75;

                if (arr.Length > 1)
                {
                    if (int.TryParse(arr[1].Replace("ms", ""), out int n))
                    {
                        if (n > 0)
                            time = n;
                    }
                }

                if (arr[0].Contains(")"))
                {
                    string[] arr2 = arr[0].Split(new string[] { ")" }, StringSplitOptions.RemoveEmptyEntries);

                    string str2 = arr2[0].Substring(5, arr2[0].Length - 5);

                    arr2 = str2.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

                    x = -1;
                    y = -1;

                    int.TryParse(arr2[0], out int n1);
                    int.TryParse(arr2[1], out int n2);

                    x = n1;
                    y = n2;
                }
                else
                {
                    x = -1;
                    y = -1;
                }
                
                if (arr[0].Contains(" while pressing "))
                {
                    string[] arr2 = arr[0].Split(new string[] { " while pressing " },
                        StringSplitOptions.RemoveEmptyEntries);

                    keys = arr2[1].Split(new string[] { " + " }, StringSplitOptions.RemoveEmptyEntries);
                }
                else
                {
                    keys = null;
                }
            }
            else if (action_text.StartsWith("Toggle: LMB")
                || action_text.StartsWith("Toggle: RMB"))
            {
                //Toggle: LMB (x, y) + alt + control

                option = 1;

                string str = action_text.Replace("Toggle: ", "");

                left = true;

                if (str.StartsWith("RMB"))
                    left = false;

                if (str.Contains(")"))
                {
                    string[] arr2 = str.Split(new string[] { ")" }, StringSplitOptions.RemoveEmptyEntries);

                    string str2 = arr2[0].Substring(5, arr2[0].Length - 5);

                    arr2 = str2.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

                    x = -1;
                    y = -1;

                    int.TryParse(arr2[0], out int n1);
                    int.TryParse(arr2[1], out int n2);

                    x = n1;
                    y = n2;
                }
                else
                {
                    x = -1;
                    y = -1;
                }

                if (str.Contains(" + "))
                {
                    string[] arr2 = str.Split(new string[] { " + " },
                        StringSplitOptions.RemoveEmptyEntries);

                    keys = new string[arr2.Length - 1];

                    for (int i = 1; i < arr2.Length; i++)
                    {
                        keys[i - 1] = arr2[i];
                    }
                }
                else
                {
                    keys = null;
                }
            }
            else if (action_text.StartsWith("Hold down: LMB")
                || action_text.StartsWith("Hold down: RMB"))
            {
                //Hold down: LMB (x, y) + alt + control
                
                option = 2;

                string str = action_text.Replace("Hold down: ", "");

                left = true;

                if (str.StartsWith("RMB"))
                    left = false;

                if (str.Contains(")"))
                {
                    string[] arr2 = str.Split(new string[] { ")" }, StringSplitOptions.RemoveEmptyEntries);

                    string str2 = arr2[0].Substring(5, arr2[0].Length - 5);

                    arr2 = str2.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

                    x = -1;
                    y = -1;

                    int.TryParse(arr2[0], out int n1);
                    int.TryParse(arr2[1], out int n2);

                    x = n1;
                    y = n2;
                }
                else
                {
                    x = -1;
                    y = -1;
                }

                if (str.Contains(" + "))
                {
                    string[] arr2 = str.Split(new string[] { " + " },
                        StringSplitOptions.RemoveEmptyEntries);

                    keys = new string[arr2.Length - 1];

                    for (int i = 1; i < arr2.Length; i++)
                    {
                        keys[i - 1] = arr2[i];
                    }
                }
                else
                {
                    keys = null;
                }
            }
            else if (action_text.StartsWith("Release: LMB")
                || action_text.StartsWith("Release: RMB"))
            {
                //Release: LMB (x, y) + alt + control

                option = 3;

                string str = action_text.Replace("Release: ", "");

                left = true;

                if (str.StartsWith("RMB"))
                    left = false;

                if (str.Contains(")"))
                {
                    string[] arr2 = str.Split(new string[] { ")" }, StringSplitOptions.RemoveEmptyEntries);

                    string str2 = arr2[0].Substring(5, arr2[0].Length - 5);

                    arr2 = str2.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

                    x = -1;
                    y = -1;

                    int.TryParse(arr2[0], out int n1);
                    int.TryParse(arr2[1], out int n2);

                    x = n1;
                    y = n2;
                }
                else
                {
                    x = -1;
                    y = -1;
                }

                if (str.Contains(" + "))
                {
                    string[] arr2 = str.Split(new string[] { " + " },
                        StringSplitOptions.RemoveEmptyEntries);

                    keys = new string[arr2.Length - 1];

                    for (int i = 1; i < arr2.Length; i++)
                    {
                        keys[i - 1] = arr2[i];
                    }
                }
                else
                {
                    keys = null;
                }
            }
        }
    }

    public class ActionMoveMouse
    {
        public bool absolute { get; set; }
        public string[] keys { get; set; }
        public int x { get; set; }
        public int y { get; set; }
        public string action_text { get; set; }

        public ActionMoveMouse()
        {
        }

        public ActionMoveMouse(string Action_text)
        {
            //Move cursor to: (x, y) while pressing alt + control
            //Move cursor by: (x, y) while pressing alt + control

            action_text = Action_text;

            string str = action_text.Replace("Move cursor ", "");

            absolute = true;

            if (str.StartsWith("by"))
                absolute = false;

            if (str.Contains(")"))
            {
                string[] arr2 = str.Split(new string[] { ")" }, StringSplitOptions.RemoveEmptyEntries);

                string str2 = arr2[0].Substring(5, arr2[0].Length - 5);

                arr2 = str2.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

                x = -1;
                y = -1;

                int.TryParse(arr2[0], out int n1);
                int.TryParse(arr2[1], out int n2);

                x = n1;
                y = n2;
            }
            else
            {
                x = -1;
                y = -1;
            }

            if (str.Contains(" while pressing "))
            {
                string[] arr2 = str.Split(new string[] { " while pressing " },
                    StringSplitOptions.RemoveEmptyEntries);

                keys = arr2[1].Split(new string[] { " + " }, StringSplitOptions.RemoveEmptyEntries);
            }
            else
            {
                keys = null;
            }
        }
    }

    public class ActionScrollMouse
    {
        public bool up { get; set; }
        public int scrolling_value { get; set; }
        public string[] keys { get; set; }
        public string action_text { get; set; }

        public ActionScrollMouse()
        {
        }

        public ActionScrollMouse(string Action_text)
        {
            //Scroll up: 40 while pressing alt + control

            action_text = Action_text;

            string str = action_text.Replace("Scroll ", "");

            up = true;

            if (str.StartsWith("down"))
                up = false;

            if (str.Contains(": "))
            {
                string[] arr = str.Split(new string[] { ": " }, StringSplitOptions.RemoveEmptyEntries);

                string[] arr2 = arr[1].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                scrolling_value = 4;

                if(int.TryParse(arr2[0], out int n1))
                {
                    scrolling_value = n1;
                }
            }

            if (str.Contains(" while pressing "))
            {
                string[] arr2 = str.Split(new string[] { " while pressing " },
                    StringSplitOptions.RemoveEmptyEntries);

                keys = arr2[1].Split(new string[] { " + " }, StringSplitOptions.RemoveEmptyEntries);
            }
            else
            {
                keys = null;
            }
        }
    }

    public class ActionOpenFileProgram
    {
        public string path;

        public ActionOpenFileProgram(string Action_text)
        {
            path = Action_text.Replace("Open: ", "");
        }
    }

    public class ActionOpenURL
    {
        public string url { get; set; }
        public string action_text { get; set; }

        public ActionOpenURL()
        {
        }

        public ActionOpenURL(string Action_text)
        {
            action_text = Action_text;

            url = action_text.Replace("Open URL: ", "");
        }
    }

    public class ActionPlaySound
    {
        public string path;

        public ActionPlaySound(string Action_text)
        {
            path = Action_text.Replace("Play sound: ", "");
        }
    }

    public class ActionReadAloud
    {
        public string text { get; set; }
        public string action_text { get; set; }

        public ActionReadAloud()
        {
        }

        public ActionReadAloud(string Action_text)
        {
            action_text = Action_text;

            text = action_text.Replace("Read aloud: ", "");
        }
    }

    public class ActionTypeText
    {
        public string text { get; set; }
        public string action_text { get; set; }

        public ActionTypeText()
        {
        }

        public ActionTypeText(string Action_text)
        {
            action_text = Action_text;

            text = action_text.Replace("Type: ", "");
        }
    }
    

    public class ActionWait
    {
        public int time { get; set; }
        public string action_text { get; set; }

        public ActionWait()
        {
        }

        public ActionWait(string Action_text)
        {
            action_text = Action_text;

            time = 1000;

            string str = action_text.Replace("Wait: ", "");
            str = str.Replace("ms", "");

            if (int.TryParse(str, out int n))
            {
                time = n;
            }
        }
    }

    public partial class MainWindow
    {
        public CollectionView cv_LVcommands;
        public CollectionView cv_LVprofiles;

        void refresh_and_save_all(string selected_command_name, int selected_profile_ind = -1)
        {
            string selected_profile_name = null;

            if(selected_profile_ind >= 0)
                selected_profile_name = Middle_Man.profiles[selected_profile_ind].name;

            Middle_Man.sort_profiles_by_name_asc();

            if(selected_profile_ind != -1)
                Middle_Man.sort_commands_by_name_asc(selected_profile_ind);

            Middle_Man.save_profiles();

            cv_LVprofiles.Refresh();

            if(cv_LVcommands != null)
                cv_LVcommands.Refresh();

            LVactions.ItemsSource = null;
            
            if (string.IsNullOrEmpty(selected_profile_name) == false)
            {
                for (int i = 0; i < LVprofiles.Items.Count; i++)
                {
                    if (((Profile)LVprofiles.Items[i]).name == selected_profile_name)
                    {
                        LVprofiles.SelectedIndex = i;
                        break;
                    }
                }
            }
            if (string.IsNullOrEmpty(selected_command_name) == false)
            {
                for (int i = 0; i < LVcommands.Items.Count; i++)
                {
                    if (((CustomCommand)LVcommands.Items[i]).name == selected_command_name)
                    {
                        LVcommands.SelectedIndex = i;
                        break;
                    }
                }
            }

            Middle_Man.force_updating_both_cc_grammars = true;
        }

        private void Badd_profile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WindowAddEditProfile w = new WindowAddEditProfile();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bedit_profile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;

                if (ind != -1)
                {
                    Profile p = ((Profile)LVprofiles.Items[ind]);

                    WindowAddEditProfile w = new WindowAddEditProfile(p.name, p.program, p.enabled);
                    w.Owner = Application.Current.MainWindow;
                    w.ShowInTaskbar = false;
                    w.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdelete_profile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;

                if (ind != -1)
                {
                    MessageBoxResult dialogResult;

                    if (LVprofiles.SelectedItems.Count == 1)
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you " +
                            "want to permanently delete the selected profile?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }
                    else
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you " +
                            "want to permanently delete the selected profiles?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }

                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        foreach (Profile p in LVprofiles.SelectedItems)
                        {
                            int index = LVprofiles.Items.IndexOf(p);

                            Middle_Man.profiles.RemoveAt(index);
                            File.Delete(Path.Combine(Middle_Man.profiles_path, p.name + ".xml"));
                        }
                    }

                    refresh_and_save_all(null);

                    LVcommands.ItemsSource = null;

                    LVprofiles.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Badd_command_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;

                if (ind != -1)
                {
                    foreach (System.Windows.Window window in Application.Current.Windows)
                    {
                        if (window.GetType() == typeof(WindowAddEditCommand))
                        {
                            throw new Exception("You may add/edit only one custom command at a time.");
                        }
                    }

                    WindowAddEditCommand w;

                    w = new WindowAddEditCommand(
                            ((Profile)LVprofiles.SelectedItems[0]).name, "", false, 0);

                    w.Owner = Application.Current.MainWindow;
                    w.ShowInTaskbar = false;
                    w.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void edit_command(int selected_action_ind = -1)
        {
            try
            {
                int ind1 = LVprofiles.SelectedIndex;
                int ind2 = LVcommands.SelectedIndex;

                if (ind1 != -1 && ind2 != -1)
                {
                    //not needed anymore, because I changed w.ShowDialog(); to w.ShowDialog() to block
                    //windows below latest window opened
                    foreach (System.Windows.Window window in Application.Current.Windows)
                    {
                        if (window.GetType() == typeof(WindowAddEditCommand))
                        {
                            throw new Exception("You may add/edit only one custom command at a time.");
                        }
                    }

                    WindowAddEditCommand w;

                    w = new WindowAddEditCommand(((Profile)LVprofiles.SelectedItems[0]).name,
                        ((CustomCommand)LVcommands.SelectedItems[0]).name, false, 0,
                        selected_action_ind);

                    w.Owner = Application.Current.MainWindow;
                    w.ShowInTaskbar = false;
                    w.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bedit_command_Click(object sender, RoutedEventArgs e)
        {
            
            edit_command();
        }

        private void Bdelete_command_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind1 = LVprofiles.SelectedIndex;
                int ind2 = LVcommands.SelectedIndex;

                if (ind1 != -1 && ind2 != -1)
                {
                    MessageBoxResult dialogResult;

                    if (LVcommands.SelectedItems.Count == 1)
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you want" +
                            " to permanently delete the selected command?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }
                    else
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you want" +
                            " to permanently delete the selected commands?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }

                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        foreach (CustomCommand cc in LVcommands.SelectedItems)
                        {
                            for (int i = 0; i < Middle_Man.profiles[ind1].custom_commands.Count; i++)
                            {
                                if (cc.name == Middle_Man.profiles[ind1].custom_commands[i].name)
                                {
                                    Middle_Man.profiles[ind1].custom_commands.RemoveAt(i);
                                    break;
                                }
                            }
                        }

                        refresh_and_save_all(null, ind1);

                        LVcommands.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bedit_action_Click(object sender, RoutedEventArgs e)
        {
            try
            { 
                int ind = LVactions.SelectedIndex;

                if (ind != -1)
                    edit_command(ind);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC027", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bexport_profiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;

                if (ind != -1)
                {
                    Microsoft.Win32.RegistryKey reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                    .OpenSubKey(Middle_Man.registry_path_easy, true);

                    string export_path = null;

                    if (reg_key_easy != null)
                    {
                        object ob = reg_key_easy.GetValue(Middle_Man.registry_key_last_export_path);

                        if (ob != null)
                        {
                            export_path = ob.ToString();
                        }
                    }

                    CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                    dialog.IsFolderPicker = true;

                    if (export_path == null)
                        dialog.InitialDirectory = Environment.GetFolderPath(
                            Environment.SpecialFolder.MyDocuments);
                    else
                        dialog.InitialDirectory = export_path;

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        if (reg_key_easy == null)
                        {
                            reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                    .CreateSubKey(Middle_Man.registry_path_easy, true);
                        }

                        reg_key_easy.SetValue(Middle_Man.registry_key_last_export_path, dialog.FileName);

                        List<Profile> chosen_profiles = new List<Profile>();

                        foreach (Profile p in LVprofiles.SelectedItems)
                        {
                            for (int i = 0; i < Middle_Man.profiles.Count; i++)
                            {
                                if (Middle_Man.profiles[i].name == p.name)
                                {
                                    chosen_profiles.Add(Middle_Man.profiles[i]);
                                }
                            }
                        }

                        Middle_Man.save_profiles(chosen_profiles, dialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC007", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bimport_profiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.RegistryKey reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                    .OpenSubKey(Middle_Man.registry_path_easy, true);

                string import_path = null;

                if (reg_key_easy != null)
                {
                    object ob = reg_key_easy.GetValue(Middle_Man.registry_key_last_import_path);

                    if (ob != null)
                    {
                        import_path = ob.ToString();
                    }
                }

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = "Import Profiles";
                ofd.Multiselect = true;
                ofd.Filter = "XML files (*.xml)|*.xml";

                if (import_path == null)
                    ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                else
                    ofd.InitialDirectory = import_path;

                if (ofd.ShowDialog() == true)
                {
                    int ind1 = LVprofiles.SelectedIndex;
                    int ind2 = LVcommands.SelectedIndex;

                    string selected_profile_name = "", selected_command_name = "";
                    string last_imported_profile_name = "";

                    if (ind1 != -1)
                        selected_profile_name = Middle_Man.profiles[ind1].name;
                    if (ind2 != -1)
                        selected_command_name = ((CustomCommand)LVcommands.SelectedItem).name;

                    if (reg_key_easy == null)
                    {
                        reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                .CreateSubKey(Middle_Man.registry_path_easy, true);
                    }

                    reg_key_easy.SetValue(Middle_Man.registry_key_last_import_path, ofd.FileName);

                    int cancelations = 0;

                    foreach (string file_path in ofd.FileNames)
                    {
                        string[] arr = file_path.Split(new string[] { "\\" },
                            StringSplitOptions.RemoveEmptyEntries);

                        string filename = arr[arr.Length - 1];
                        string profile_name = filename;

                        if (profile_name.Contains("."))
                            profile_name = filename.Substring(0, filename.Length - 4);

                        bool found = false;

                        for (int i = 0; i < Middle_Man.profiles.Count; i++)
                        {
                            if (Middle_Man.profiles[i].name == profile_name)
                            {
                                found = true;

                                MessageBoxResult dialogResult = System.Windows.MessageBox.Show(
                                    "There is already a profile named \"" + profile_name + "\"."
                                    + " Do you want to replace it?",
                                      "Profile Already Exists", MessageBoxButton.YesNo,
                                      MessageBoxImage.Question);

                                if (dialogResult == MessageBoxResult.Yes)
                                {
                                    Middle_Man.profiles.RemoveAt(i);
                                    i--;

                                    load_profile(file_path);
                                    last_imported_profile_name = profile_name;
                                }
                                else
                                    cancelations++;

                                break;
                            }
                        }

                        if (found == false)
                        {
                            load_profile(file_path);
                            last_imported_profile_name = profile_name;
                        }
                    }

                    if (cancelations < ofd.FileNames.Length)
                    {
                        refresh_and_save_all(selected_command_name, ind1);
                    }

                    if (string.IsNullOrEmpty(selected_profile_name))
                    {
                        int s_ind = Middle_Man.get_profile_ind_by_name(selected_profile_name);
                        LVprofiles.SelectedIndex = s_ind;
                    }
                    else if (string.IsNullOrEmpty(last_imported_profile_name))
                    {
                        int s_ind = Middle_Man.get_profile_ind_by_name(last_imported_profile_name);
                        LVprofiles.SelectedIndex = s_ind;
                    }

                    LVprofiles.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC008", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVprofiles_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;

                if (ind != -1)
                {
                    Bedit_profile.IsEnabled = true;
                    Bdelete_profile.IsEnabled = true;
                    Bexport_profiles.IsEnabled = true;
                    Badd_command.IsEnabled = true;

                    MIenable_profiles.IsEnabled = true;
                    MIdisable_profiles.IsEnabled = true;
                    MIedit_profiles.IsEnabled = true;
                    MIdelete_profiles.IsEnabled = true;
                    MIduplicate_profiles.IsEnabled = true;

                    LVactions.ItemsSource = null;

                    LVcommands.ItemsSource = Middle_Man.profiles[ind].custom_commands;

                    cv_LVcommands = (CollectionView)CollectionViewSource.GetDefaultView(
                                LVcommands.ItemsSource);

                    if (cv_LVcommands.GroupDescriptions.Count == 0)
                    {
                        PropertyGroupDescription pgd = new PropertyGroupDescription("group");
                        cv_LVcommands.GroupDescriptions.Add(pgd);
                        cv_LVcommands.SortDescriptions.Add(new SortDescription("group",
                            ListSortDirection.Ascending));
                        cv_LVcommands.SortDescriptions.Add(new SortDescription("name",
                            ListSortDirection.Ascending));
                    }
                }
                else
                {
                    Bedit_profile.IsEnabled = false;
                    Bdelete_profile.IsEnabled = false;
                    Bexport_profiles.IsEnabled = false;
                    Badd_command.IsEnabled = false;

                    MIenable_profiles.IsEnabled = false;
                    MIdisable_profiles.IsEnabled = false;
                    MIedit_profiles.IsEnabled = false;
                    MIdelete_profiles.IsEnabled = false;
                    MIduplicate_profiles.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC009", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVcommands_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                int ind1 = LVprofiles.SelectedIndex;
                int ind2 = LVcommands.SelectedIndex;

                if (ind1 != -1 && ind2 != -1)
                {
                    Bedit_command.IsEnabled = true;
                    Bdelete_command.IsEnabled = true;

                    MIenable_commands.IsEnabled = true;
                    MIdisable_commands.IsEnabled = true;
                    MIedit_commands.IsEnabled = true;
                    MIdelete_commands.IsEnabled = true;
                    MIcopy_commands.IsEnabled = true;
                    MIpaste_commands.IsEnabled = true;

                    int com_ind = Middle_Man.get_command_ind_by_name(ind1,
                        ((CustomCommand)LVcommands.SelectedItem).name);

                    LVactions.ItemsSource = Middle_Man.profiles[ind1].custom_commands[com_ind].actions;

                    CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(
                                LVactions.ItemsSource);
                    cv.Refresh();
                }
                else
                {
                    Bedit_command.IsEnabled = false;
                    Bdelete_command.IsEnabled = false;

                    MIenable_commands.IsEnabled = false;
                    MIdisable_commands.IsEnabled = false;
                    MIedit_commands.IsEnabled = false;
                    MIdelete_commands.IsEnabled = false;
                    MIcopy_commands.IsEnabled = false;
                    MIpaste_commands.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC010", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVactions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                int ind1 = LVprofiles.SelectedIndex;
                int ind2 = LVcommands.SelectedIndex;
                int ind3 = LVactions.SelectedIndex;

                if (ind1 != -1 && ind2 != -1 && ind3 != -1)
                {
                    Bedit_action.IsEnabled = true;
                    MIedit_actions.IsEnabled = true;
                }
                else
                {
                    Bedit_action.IsEnabled = false;
                    MIedit_actions.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC027", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVprofiles_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Delete)
                {
                    Bdelete_profile_Click(null, null);
                }
                else if (e.Key == Key.V
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIpaste_commands_Click(null, null);
                }
                else if (e.Key == Key.D
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIduplicate_profiles_Click(null, null);
                }
                else if (e.Key == Key.OemPlus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIenable_profiles_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIdisable_profiles_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC011", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVcommands_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            { 
                if (e.Key == Key.Delete)
                {
                    Bdelete_command_Click(null, null);
                }
                else if (e.Key == Key.C 
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIcopy_commands_Click(null, null);
                }
                else if (e.Key == Key.V
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIpaste_commands_Click(null, null);
                }
                else if (e.Key == Key.OemPlus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIenable_commands_Click(null, null);
                }
                else if (e.Key == Key.OemMinus
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIdisable_commands_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC012", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIenable_profiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVprofiles.SelectedIndex != -1)
                {
                    foreach (Profile p in LVprofiles.SelectedItems)
                    {
                        for (int i = 0; i < Middle_Man.profiles.Count; i++)
                        {
                            if (Middle_Man.profiles[i].name == p.name)
                            {
                                Middle_Man.profiles[i].enabled = true;
                                break;
                            }
                        }
                    }

                    Middle_Man.save_profiles();

                    cv_LVprofiles.Refresh();

                    Middle_Man.force_updating_both_cc_grammars = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC024", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIdisable_profiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVprofiles.SelectedIndex != -1)
                {
                    foreach (Profile p in LVprofiles.SelectedItems)
                    {
                        for (int i = 0; i < Middle_Man.profiles.Count; i++)
                        {
                            if (Middle_Man.profiles[i].name == p.name)
                            {
                                Middle_Man.profiles[i].enabled = false;
                                break;
                            }
                        }
                    }

                    Middle_Man.save_profiles();

                    cv_LVprofiles.Refresh();

                    Middle_Man.force_updating_both_cc_grammars = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC025", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIdelete_profiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if(LVprofiles.SelectedIndex != -1)
                    Bdelete_profile_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC013", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIduplicate_profiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;

                if (ind != -1)
                {
                    foreach (Profile p in LVprofiles.SelectedItems)
                    {
                        for (int i = 0; i < Middle_Man.profiles.Count; i++)
                        {
                            if (Middle_Man.profiles[i].name == p.name)
                            {
                                int j = 1;
                                string name = p.name;
                                string new_name = name + " copy " + j;

                                bool found = false;

                                do
                                {
                                    found = false;

                                    foreach (Profile p2 in Middle_Man.profiles)
                                    {
                                        if (p2.name == new_name)
                                        {
                                            found = true;
                                            j++;
                                            new_name = name + " copy " + j;

                                            break;
                                        }
                                    }
                                }
                                while (found);

                                //We need to use DeepCopy, because we don't want to assign a reference to
                                //the list like this: custom_commands = Middle_Man.profiles[i].custom_commands

                                List<CustomCommand> new_list = new List<CustomCommand>();

                                foreach(CustomCommand cc in Middle_Man.profiles[i].custom_commands)
                                {
                                    new_list.Add(new CustomCommand()
                                    {
                                        actions = DeepCopy(cc.actions),
                                        description = cc.description,
                                        enabled = cc.enabled,
                                        max_executions = cc.max_executions,
                                        group = cc.group,
                                        name = cc.name
                                    });
                                }

                                Middle_Man.profiles.Add(new Profile()
                                {
                                    name = new_name,
                                    enabled = Middle_Man.profiles[i].enabled,
                                    program = Middle_Man.profiles[i].program,
                                    custom_commands = new_list
                                });

                                /* 
                                //Not working properly when obfuscated by obfuscar (DeepCopying custom
                                //commands is the cause of the problem, but DeepCopying actions works):

                                Middle_Man.profiles.Add(new Profile()
                                {
                                    name = new_name,
                                    enabled = Middle_Man.profiles[i].enabled,
                                    program = Middle_Man.profiles[i].program,
                                    custom_commands = DeepCopy(Middle_Man.profiles[i].custom_commands)
                                });
                                */

                                break;
                            }
                        }
                    }

                    Middle_Man.sort_profiles_by_name_asc();

                    Middle_Man.save_profiles();

                    cv_LVprofiles.Refresh();

                    Middle_Man.force_updating_both_cc_grammars = true;
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.StartsWith("The trial version allows"))
                    MessageBox.Show(ex.Message, "Trial version limit reached",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                else
                    MessageBox.Show(ex.Message, "Error CC014", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIedit_profiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if(LVprofiles.SelectedIndex != -1)
                    Bedit_profile_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC023", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIdisable_commands_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;
                int ind_com = LVcommands.SelectedIndex;

                if (ind != -1 && ind_com != -1)
                {
                    foreach (CustomCommand cc in LVcommands.SelectedItems)
                    {
                        for (int i = 0; i < Middle_Man.profiles[ind].custom_commands.Count; i++)
                        {
                            if (cc.name == Middle_Man.profiles[ind].custom_commands[i].name)
                            {
                                Middle_Man.profiles[ind].custom_commands[i].enabled = false;
                                break;
                            }
                        }
                    }

                    Middle_Man.save_profiles();

                    cv_LVcommands.Refresh();

                    Middle_Man.force_updating_both_cc_grammars = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC015", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIenable_commands_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;
                int ind_com = LVcommands.SelectedIndex;

                if (ind != -1 && ind_com != -1)
                {
                    foreach (CustomCommand cc in LVcommands.SelectedItems)
                    {
                        for (int i = 0; i < Middle_Man.profiles[ind].custom_commands.Count; i++)
                        {
                            if (cc.name == Middle_Man.profiles[ind].custom_commands[i].name)
                            {
                                Middle_Man.profiles[ind].custom_commands[i].enabled = true;
                                break;
                            }
                        }
                    }

                    Middle_Man.save_profiles();

                    cv_LVcommands.Refresh();

                    Middle_Man.force_updating_both_cc_grammars = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC016", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIedit_commands_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVcommands.SelectedIndex != -1)
                    edit_command();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC017", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIdelete_commands_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVcommands.SelectedIndex != -1)
                    Bdelete_command_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC018", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIcopy_commands_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;
                int ind_com = LVcommands.SelectedIndex;

                if (ind != -1 && ind_com != -1)
                {
                    Middle_Man.copied_commands = new List<CustomCommand>();

                    foreach (CustomCommand cc in LVcommands.SelectedItems)
                    {
                        for (int i = 0; i < Middle_Man.profiles[ind].custom_commands.Count; i++)
                        {
                            if (cc.name == Middle_Man.profiles[ind].custom_commands[i].name)
                            {
                                Middle_Man.copied_commands.Add(Middle_Man.profiles[ind].custom_commands[i]);
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC019", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIpaste_commands_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVprofiles.SelectedIndex;
                
                if (ind != -1)
                {
                    foreach (CustomCommand cc in Middle_Man.copied_commands)
                    {
                        bool found2 = false;

                        for (int i = 0; i < Middle_Man.profiles[ind].custom_commands.Count; i++)
                        {
                            if (Middle_Man.profiles[ind].custom_commands[i].name == cc.name)
                            {
                                found2 = true;

                                int j = 1;
                                string name = cc.name;
                                string new_name = name + " copy " + j;

                                bool found = false;

                                do
                                {
                                    found = false;

                                    foreach (CustomCommand cc2 in Middle_Man.profiles[ind].custom_commands)
                                    {
                                        if (cc2.name == new_name)
                                        {
                                            found = true;
                                            j++;
                                            new_name = name + " copy " + j;

                                            break;
                                        }
                                    }
                                }
                                while (found);

                                //new CustomComman and DeepCopy is required here,
                                //because we don't want to reference an object
                                Middle_Man.profiles[ind].custom_commands.Add(new CustomCommand()
                                {
                                    name = new_name,
                                    enabled = cc.enabled,
                                    description = cc.description,
                                    group = cc.group,
                                    max_executions = cc.max_executions,
                                    actions = DeepCopy(cc.actions)
                                });

                                break;
                            }
                        }

                        if (found2 == false)
                        {
                            //new CustomComman and DeepCopy is required here,
                            //because we don't want to reference an object
                            Middle_Man.profiles[ind].custom_commands.Add(new CustomCommand()
                            {
                                name = cc.name,
                                enabled = cc.enabled,
                                description = cc.description,
                                group = cc.group,
                                max_executions = cc.max_executions,
                                actions = DeepCopy(cc.actions)
                            });
                        }
                    }

                    Middle_Man.sort_commands_by_name_asc(ind);

                    Middle_Man.save_profiles();

                    cv_LVcommands.Refresh();

                    Middle_Man.force_updating_both_cc_grammars = true;
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.StartsWith("The trial version allows"))
                    MessageBox.Show(ex.Message, "Trial version limit reached",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                else
                    MessageBox.Show(ex.Message, "Error CC020", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVprofiles_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (LVprofiles.SelectedIndex != -1)
                    Bedit_profile_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC021", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVcommands_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (LVcommands.SelectedIndex != -1)
                    Bedit_command_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC022", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIedit_actions_Click(object sender, RoutedEventArgs e)
        {
            Bedit_action_Click(null, null);
        }

        private void LVactions_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            int ind = LVactions.SelectedIndex;

            if (ind != -1)
                edit_command(ind);
        }

        void execute_custom_commands(List<CC_Action> actions)
        {
            recognition_suspended = true;

            try
            {
                if (read_recognized_speech)
                {
                    if (ss.Volume != ss_volume)
                        ss.Volume = ss_volume;

                    ss.SpeakAsync(r);
                }

                string recognized_speech = r.FirstCharToUpper();

                string[] a = r.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                int n = a.Length;
                int executions = 1;

                if (a[n - 1] == "twice" || a[n - 1] == "times")
                {
                    if (a[n - 1] == "twice")
                    {
                        executions = 2;
                        r = r.Remove(r.Length - 6, 6);
                    }
                    else
                    {
                        executions = int.Parse(a[n - 2]);
                        r = r.Remove(r.Length - 7 - a[n - 2].Length, a[n - 2].Length + 7);
                    }
                }

                for (int i = 1; i <= executions; i++)
                {
                    for(int j=0; j<actions.Count; j++)
                    { 
                        string action_text = actions[j].action;

                        //SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                        //{
                        //    SW.TBrecognized_speech.Text = recognized_speech + "\r\n" +
                        //        "Execution: " + i + "/" + executions + "\r\n" +
                        //        "Action: " + (j + 1) + "/" + actions.Count + "\r\n" +
                        //        action_text;
                        //}));

                        if (j == 0)
                        {
                            SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                SW.TBrecognized_speech.Text = recognized_speech + "\r\n" +
                                    "Execution: " + i + "/" + executions +
                                    ", Action: " + (j + 1) + "/" + actions.Count + "\r\n" +
                                    (j + 1) + ": " + action_text;
                            }));
                        }
                        else
                        {
                            string action_text_substring = actions[j - 1].action;
                            
                            action_text_substring = action_text_substring.Replace("\r\n", "");
                            action_text_substring = action_text_substring.Replace("\t", "");
                            int max_length = 36;

                            if (action_text_substring.Length > max_length)
                                action_text_substring = action_text_substring.Substring(0, max_length) 
                                    + "...";

                            SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                SW.TBrecognized_speech.Text = recognized_speech + "\r\n" +
                                    "Execution: " + i + "/" + executions +
                                    ", Action: " + (j + 1) + "/" + actions.Count + "\r\n" +
                                    j + ": " + action_text_substring + "\r\n" +
                                    (j + 1) + ": " + action_text;
                            }));
                        }

                        if (action_text.StartsWith("Click: LMB")
                            || action_text.StartsWith("Click: RMB")
                            || action_text.StartsWith("Toggle: LMB")
                            || action_text.StartsWith("Toggle: RMB")
                            || action_text.StartsWith("Hold down: LMB")
                            || action_text.StartsWith("Hold down: RMB")
                            || action_text.StartsWith("Release: LMB")
                            || action_text.StartsWith("Release: RMB"))
                        {
                            ActionMouse am = new ActionMouse(action_text);

                            List<VirtualKeyCode> keys = new List<VirtualKeyCode>();

                            if (am.option == 0) //click
                            {
                                if (am.x > -1 && am.y > -1)
                                {
                                    move_mouse(am.x, am.y);
                                }
                                else
                                {
                                    am.x = System.Windows.Forms.Cursor.Position.X;
                                    am.y = System.Windows.Forms.Cursor.Position.Y;
                                }

                                if (am.keys != null)
                                {
                                    foreach (string key in am.keys)
                                    {
                                        keys.Add(Middle_Man.get_virtual_key_code_by_key_name(key));
                                    }

                                    foreach (VirtualKeyCode vkc in keys)
                                    {
                                        key_down(vkc);
                                    }
                                }

                                if (am.left)
                                    LMBClick(am.x, am.y, am.time);
                                else
                                    RMBClick(am.x, am.y, am.time);

                                if (am.keys != null)
                                {
                                    foreach (VirtualKeyCode vkc in keys)
                                    {
                                        key_up(vkc);
                                    }
                                }
                            }
                            else if (am.option == 1) //toggle
                            {
                                if (am.x > -1 && am.y > -1)
                                {
                                    real_mouse_move(am.x, am.y);
                                    Thread.Sleep(10);
                                }

                                if (am.keys != null)
                                {
                                    List<VirtualKeyCode> down_keys = new List<VirtualKeyCode>();
                                    List<VirtualKeyCode> up_keys = new List<VirtualKeyCode>();

                                    foreach (string key in am.keys)
                                    {
                                        VirtualKeyCode vkc = Middle_Man.get_virtual_key_code_by_key_name(key);

                                        if (sim.InputDeviceState.IsKeyDown(vkc))
                                        {
                                            down_keys.Add(vkc);
                                        }
                                        else
                                        {
                                            up_keys.Add(vkc);
                                        }
                                    }

                                    if (down_keys.Count > 0)
                                        remove_keys_from_keys_to_hold(down_keys);

                                    if (up_keys.Count > 0)
                                        add_keys_to_keys_to_hold(up_keys);
                                }                                

                                if (am.left)
                                {
                                    if (keys_to_hold.Contains(VirtualKeyCode.LBUTTON))
                                        remove_keys_from_keys_to_hold(new List<VirtualKeyCode>()
                                        { VirtualKeyCode.LBUTTON });
                                    else
                                        add_keys_to_keys_to_hold(new List<VirtualKeyCode>()
                                        { VirtualKeyCode.LBUTTON });
                                }
                                else
                                {
                                    if (keys_to_hold.Contains(VirtualKeyCode.RBUTTON))
                                        remove_keys_from_keys_to_hold(new List<VirtualKeyCode>()
                                        { VirtualKeyCode.RBUTTON });
                                    else
                                        add_keys_to_keys_to_hold(new List<VirtualKeyCode>()
                                        { VirtualKeyCode.RBUTTON });
                                }
                            }
                            else if (am.option == 2) //hold down
                            {
                                if (am.x > -1 && am.y > -1)
                                {
                                    real_mouse_move(am.x, am.y);
                                    Thread.Sleep(10);
                                }

                                if (am.keys != null)
                                {
                                    foreach (string key in am.keys)
                                    {
                                        keys.Add(Middle_Man.get_virtual_key_code_by_key_name(key));
                                    }

                                    add_keys_to_keys_to_hold(keys);
                                }                                

                                if (am.left)
                                    add_keys_to_keys_to_hold(new List<VirtualKeyCode>()
                                        { VirtualKeyCode.LBUTTON });
                                else
                                    add_keys_to_keys_to_hold(new List<VirtualKeyCode>()
                                        { VirtualKeyCode.RBUTTON });
                            }
                            else if (am.option == 3) //release
                            {
                                if (am.x > -1 && am.y > -1)
                                {
                                    real_mouse_move(am.x, am.y);
                                    Thread.Sleep(10);
                                }

                                if (am.left)
                                    remove_keys_from_keys_to_hold(new List<VirtualKeyCode>()
                                        { VirtualKeyCode.LBUTTON });
                                else
                                    remove_keys_from_keys_to_hold(new List<VirtualKeyCode>()
                                        { VirtualKeyCode.RBUTTON });

                                if (am.keys != null)
                                {
                                    foreach (string key in am.keys)
                                    {
                                        keys.Add(Middle_Man.get_virtual_key_code_by_key_name(key));
                                    }

                                    remove_keys_from_keys_to_hold(keys);
                                }
                            }
                        }
                        else if (action_text.StartsWith("Press:")
                            || action_text.StartsWith("Toggle:")
                            || action_text.StartsWith("Hold down:")
                            || action_text.StartsWith("Release:"))
                        {
                            ActionKeyboard ak = new ActionKeyboard(action_text);

                            List<VirtualKeyCode> keys = new List<VirtualKeyCode>();

                            if (ak.option == 0) //press
                            {
                                foreach (string key in ak.keys)
                                {
                                    keys.Add(Middle_Man.get_virtual_key_code_by_key_name(key));
                                }

                                foreach (VirtualKeyCode vkc in keys)
                                {
                                    key_down(vkc);
                                }

                                Thread.Sleep(ak.time);

                                foreach (VirtualKeyCode vkc in keys)
                                {
                                    key_up(vkc);
                                }
                            }
                            else if (ak.option == 1) //toggle
                            {
                                List<VirtualKeyCode> down_keys = new List<VirtualKeyCode>();
                                List<VirtualKeyCode> up_keys = new List<VirtualKeyCode>();

                                foreach (string key in ak.keys)
                                {
                                    VirtualKeyCode vkc = Middle_Man.get_virtual_key_code_by_key_name(key);

                                    if (sim.InputDeviceState.IsKeyDown(vkc))
                                    {
                                        down_keys.Add(vkc);
                                    }
                                    else
                                    {
                                        up_keys.Add(vkc);
                                    }
                                }

                                if (down_keys.Count > 0)
                                    remove_keys_from_keys_to_hold(down_keys);

                                if (up_keys.Count > 0)
                                    add_keys_to_keys_to_hold(up_keys);
                            }
                            else if (ak.option == 2) //hold down
                            {
                                foreach (string key in ak.keys)
                                {
                                    keys.Add(Middle_Man.get_virtual_key_code_by_key_name(key));
                                }

                                add_keys_to_keys_to_hold(keys);
                            }
                            else if (ak.option == 3) //release
                            {
                                foreach (string key in ak.keys)
                                {
                                    keys.Add(Middle_Man.get_virtual_key_code_by_key_name(key));
                                }

                                remove_keys_from_keys_to_hold(keys);
                            }
                        }
                        else if (action_text.StartsWith("Release all keys and buttons"))
                        {
                            release_buttons_and_keys();
                        }
                        else if (action_text.StartsWith("Move cursor"))
                        {
                            ActionMoveMouse amm = new ActionMoveMouse(action_text);

                            List<VirtualKeyCode> keys = new List<VirtualKeyCode>();

                            if (amm.keys != null)
                            {
                                foreach (string key in amm.keys)
                                {
                                    keys.Add(Middle_Man.get_virtual_key_code_by_key_name(key));
                                }

                                add_keys_to_keys_to_hold(keys);
                            }

                            if (amm.absolute)
                                real_mouse_move(amm.x, amm.y);
                            else
                                real_move_mouse_by(amm.x, amm.y);

                            if (amm.keys != null)
                            {
                                remove_keys_from_keys_to_hold(keys);
                            }
                        }
                        else if (action_text.StartsWith("Scroll"))
                        {
                            ActionScrollMouse asm = new ActionScrollMouse(action_text);

                            List<VirtualKeyCode> keys = new List<VirtualKeyCode>();

                            if (asm.keys != null)
                            {
                                foreach (string key in asm.keys)
                                {
                                    keys.Add(Middle_Man.get_virtual_key_code_by_key_name(key));
                                }

                                add_keys_to_keys_to_hold(keys);
                            }

                            if (asm.up)
                                sim.Mouse.VerticalScroll(asm.scrolling_value);
                            else
                                sim.Mouse.VerticalScroll(asm.scrolling_value * -1);

                            if (asm.keys != null)
                            {
                                remove_keys_from_keys_to_hold(keys);
                            }
                        }
                        else if (action_text.StartsWith("Open:"))
                        {
                            ActionOpenFileProgram aofp = new ActionOpenFileProgram(action_text);

                            Process.Start(aofp.path);
                        }
                        else if (action_text.StartsWith("Open URL:"))
                        {
                            ActionOpenURL aou = new ActionOpenURL(action_text);

                            string url = aou.url;

                            if (url.StartsWith("https://") == false && url.StartsWith("http://") == false)
                            {
                                url = "https://" + url;
                            }

                            Process.Start(url);
                        }
                        else if (action_text.StartsWith("Play sound:"))
                        {
                            ActionPlaySound aps = new ActionPlaySound(action_text);

                            MediaPlayer mp = new MediaPlayer();

                            mp.Open(new Uri(aps.path));
                            mp.Play();
                        }
                        else if (action_text.StartsWith("Read aloud:"))
                        {
                            ActionReadAloud ara = new ActionReadAloud(action_text);

                            ss.Speak(ara.text);
                        }
                        else if (action_text.StartsWith("Type:"))
                        {
                            ActionTypeText att = new ActionTypeText(action_text);

                            //sim.Keyboard.TextEntry(att.text);

                            /*
                            string text = att.text;

                            string[] arr = text.Split(new string[] { "\r\n" },
                                StringSplitOptions.None);

                            for (int k = 0; k < arr.Length; k++)
                            {
                                if (string.IsNullOrEmpty(arr[k]) == false)
                                    sim.Keyboard.TextEntry(arr[k]);

                                if (k < arr.Length - 1)
                                    key_press(VirtualKeyCode.RETURN, 1);
                            }
                            */

                            //Clipboard.SetText(att.text); //requires STA Thread

                            Thread thread = new Thread(() => Clipboard.SetText(att.text));
                            thread.SetApartmentState(ApartmentState.STA); //Set the thread to STA
                            thread.Start();
                            thread.Join(); //Wait for the thread to end

                            key_down(VirtualKeyCode.LCONTROL);
                            key_down(VirtualKeyCode.VK_V);
                            Thread.Sleep(75); //don't decrease it, because long text require more time to
                                              //get pasted
                            key_up(VirtualKeyCode.VK_V);
                            key_up(VirtualKeyCode.LCONTROL);
                        }
                        else if (action_text.StartsWith("Wait:"))
                        {
                            ActionWait aw = new ActionWait(action_text);

                            if (aw.time > 0)
                                Thread.Sleep(aw.time);
                        }
                    }
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error CC026", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            recognition_suspended = false;
        }
    }
}