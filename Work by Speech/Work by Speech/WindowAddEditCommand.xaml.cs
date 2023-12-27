using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditCommand : Window
    {
        string cc_profile;
        string prev_name = "";
        bool edit = false;
        int selected_action_ind = -1;

        public List<CC_Action> actions = new List<CC_Action>();
        public CollectionView cv_actions;

        bool TM;
        short CCL;
        //tm - trial mode, ccl - custom commands number limit for trial mode
        public WindowAddEditCommand(string profile, string Prev_name, bool tm, short ccl,
            int Selected_action_ind = -1)
        {
            try
            {
                InitializeComponent();

                cc_profile = profile;
                prev_name = Prev_name;

                TM = tm;
                CCL = ccl;

                foreach (Group group in Middle_Man.groups)
                {
                    CBgroup.Items.Add(group.name);
                }

                TBmax_executions.Text = Middle_Man.last_used_max_executions;

                if (prev_name != "")
                {
                    this.Title = "Edit command added to profile: " + profile;

                    edit = true;

                    int profile_ind = Middle_Man.get_profile_ind_by_name(profile);
                    int ind = Middle_Man.get_command_ind_by_name(profile_ind, prev_name);
                    
                    TBname.Text = Middle_Man.profiles[profile_ind].custom_commands[ind].name;
                    TBdescription.Text = Middle_Man.profiles[profile_ind].custom_commands[ind].description;
                    CBgroup.SelectedItem = Middle_Man.profiles[profile_ind].custom_commands[ind].group;
                    TBmax_executions.Text =
                        Middle_Man.profiles[profile_ind].custom_commands[ind].max_executions.ToString();
                    CHBenabled.IsChecked = Middle_Man.profiles[profile_ind].custom_commands[ind].enabled;

                    foreach (CC_Action action in Middle_Man.profiles[profile_ind].custom_commands[ind].actions)
                    {
                        actions.Add(action);
                    }
                }
                else
                {
                    this.Title = "Add command to profile: " + profile;

                    int profile_ind = Middle_Man.get_profile_ind_by_name(profile);

                    CBgroup.SelectedItem = "General";

                    foreach (System.Windows.Window window in Application.Current.Windows)
                    {
                        if (window.GetType() == typeof(MainWindow))
                        {
                            MainWindow w = (MainWindow)window;

                            int ind = w.LVcommands.SelectedIndex;

                            if (ind != -1)
                                CBgroup.SelectedItem = 
                                    Middle_Man.profiles[profile_ind].custom_commands[ind].group;
                        }
                    }
                }

                LVactions.ItemsSource = actions;

                cv_actions = (CollectionView)CollectionViewSource.GetDefaultView(LVactions.Items);

                LVnew_actions.ItemsSource = Middle_Man.action_types;

                LVnew_actions.SelectedIndex = 0;

                Bedit.IsEnabled = false;
                Bcopy.IsEnabled = false;
                if (Middle_Man.copied_actions.Count == 0)
                {
                    Bpaste.IsEnabled = false;
                    MIpaste_actions.IsEnabled = false;
                }
                Bdelete.IsEnabled = false;
                Bmove_up.IsEnabled = false;
                Bmove_down.IsEnabled = false;

                MIedit_actions.IsEnabled = false;
                MIcopy_actions.IsEnabled = false;
                MIdelete_actions.IsEnabled = false;
                MImove_up_actions.IsEnabled = false;
                MImove_down_actions.IsEnabled = false;

                selected_action_ind = Selected_action_ind;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                TBname.Focus();
                TBname.CaretIndex = TBname.Text.Length;

                if (selected_action_ind >= 0)
                {
                    LVactions.SelectedIndex = selected_action_ind;

                    LVactions.ScrollIntoView(LVactions.SelectedItem);
                    LVactions.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bmanage_groups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WindowManageGroups w = new WindowManageGroups();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bedit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVactions.SelectedIndex;

                if (ind != -1)
                {
                    string action_text = ((CC_Action)LVactions.Items[ind]).action;

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

                        WindowAddEditActionMouse w = new WindowAddEditActionMouse(am, ind);
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text.StartsWith("Press:")
                        || action_text.StartsWith("Toggle:")
                        || action_text.StartsWith("Hold down:")
                        || action_text.StartsWith("Release:"))
                    {
                        ActionKeyboard ak = new ActionKeyboard(action_text);

                        WindowAddEditActionKeyboard w = new WindowAddEditActionKeyboard(ak, ind);
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text.StartsWith("Move cursor"))
                    {
                        ActionMoveMouse amm = new ActionMoveMouse(action_text);

                        WindowAddEditActionMoveMouse w = new WindowAddEditActionMoveMouse(amm, ind);
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text.StartsWith("Scroll"))
                    {
                        ActionScrollMouse asm = new ActionScrollMouse(action_text);

                        WindowAddEditActionScrollMouse w = new WindowAddEditActionScrollMouse(asm, ind);
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text.StartsWith("Open:"))
                    {
                        ActionOpenFileProgram aofp = new ActionOpenFileProgram(action_text);

                        OpenFileDialog ofd = new OpenFileDialog();
                        ofd.Title = "Choose a file/program";
                        ofd.Multiselect = false;
                        ofd.Filter = "All files (*.*)|*.*";
                        ofd.InitialDirectory = aofp.path;

                        if (ofd.ShowDialog() == true)
                        {
                            actions[ind].action = "Open: " + ofd.FileName;

                            cv_actions.Refresh();

                            LVactions.SelectedIndex = ind;

                            LVactions.ScrollIntoView(LVactions.SelectedItem);
                        }
                    }
                    else if (action_text.StartsWith("Open URL:"))
                    {
                        ActionOpenURL aou = new ActionOpenURL(action_text);

                        WindowAddEditActionOpenURL w = new WindowAddEditActionOpenURL(aou, ind);
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text.StartsWith("Play sound:"))
                    {
                        ActionPlaySound aps = new ActionPlaySound(action_text);

                        OpenFileDialog ofd = new OpenFileDialog();
                        ofd.Title = "Choose an mp3 file";
                        ofd.Multiselect = false;
                        ofd.Filter = "MP3 files (*.mp3)|*.mp3|All files (*.*)|*.*";
                        ofd.InitialDirectory = aps.path;

                        if (ofd.ShowDialog() == true)
                        {
                            actions[ind].action = "Play sound: " + ofd.FileName;

                            cv_actions.Refresh();

                            LVactions.SelectedIndex = ind;

                            LVactions.ScrollIntoView(LVactions.SelectedItem);
                        }
                    }
                    else if (action_text.StartsWith("Read aloud:"))
                    {
                        ActionReadAloud ara = new ActionReadAloud(action_text);

                        WindowAddEditActionReadText w = new WindowAddEditActionReadText(ara, ind);
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text.StartsWith("Type:"))
                    {
                        ActionTypeText att = new ActionTypeText(action_text);

                        WindowAddEditActionTypeText w = new WindowAddEditActionTypeText(att, ind);
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text.StartsWith("Wait:"))
                    {
                        ActionWait aw = new ActionWait(action_text);

                        WindowAddEditActionWait w = new WindowAddEditActionWait(aw, ind);
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bcopy_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVactions.SelectedIndex;

                Middle_Man.copied_actions = new List<CC_Action>();

                if (ind != -1)
                {
                    int ind_first = LVactions.Items.IndexOf(LVactions.SelectedItems[0]);
                    int ind_last = LVactions.Items.IndexOf(
                        LVactions.SelectedItems[LVactions.SelectedItems.Count - 1]);

                    //when user selects from top to bottom
                    if (ind_first <= ind_last)
                    {
                        foreach (CC_Action action in LVactions.SelectedItems)
                        {
                            //new CC_Action() is required here, because we don't want to reference an object
                            Middle_Man.copied_actions.Add(new CC_Action()
                            {
                                action = action.action
                            });
                        }
                    }
                    //when user selects from bottom to top
                    else
                    {
                        for (int i = LVactions.SelectedItems.Count - 1; i >= 0; i--)
                        {
                            int ind2 = LVactions.Items.IndexOf(LVactions.SelectedItems[i]);

                            //new CC_Action() is required here, because we don't want to reference an object
                            Middle_Man.copied_actions.Add(new CC_Action()
                            {
                                action = actions[ind2].action
                            });
                        }
                    }
                    
                    Bpaste.IsEnabled = true;
                    MIpaste_actions.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bpaste_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Middle_Man.copied_actions.Count > 0)
                {
                    int ind = LVactions.SelectedIndex;

                    if (ind != -1)
                    {
                        ind++;
                        foreach (CC_Action action in Middle_Man.copied_actions)
                        {
                            //new used, because we don't want a reference
                            actions.Insert(ind, new CC_Action()
                            {
                                action = action.action
                            });
                            ind++;
                        }
                    }
                    else
                    {
                        foreach (CC_Action action in Middle_Man.copied_actions)
                        {
                            //new used, because we don't want a reference
                            actions.Add(new CC_Action()
                            {
                                action = action.action
                            });
                            ind++;
                        }
                    }

                    cv_actions.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVactions.SelectedIndex;

                if (ind != -1)
                {
                    MessageBoxResult dialogResult;

                    if (LVactions.SelectedItems.Count == 1)
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you " +
                            "want to delete the selected action?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }
                    else
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you " +
                            "want to delete the selected actions?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }

                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        foreach (CC_Action action in LVactions.SelectedItems)
                        {
                            int index = LVactions.Items.IndexOf(action);

                            actions.RemoveAt(index);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC007", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cv_actions.Refresh();
            }
        }

        private void Bmove_up_Click(object sender, RoutedEventArgs e)
        {
            List<int> indexes = new List<int>();

            try
            {
                if (LVactions.SelectedIndex != -1)
                {
                    int ind_first = LVactions.Items.IndexOf(LVactions.SelectedItems[0]);
                    int ind_last = LVactions.Items.IndexOf(
                        LVactions.SelectedItems[LVactions.SelectedItems.Count - 1]);

                    //when user select from top to bottom
                    if (ind_first <= ind_last)
                    {
                        foreach (CC_Action action in LVactions.SelectedItems)
                        {
                            int ind = LVactions.Items.IndexOf(action);

                            if (ind != 0)
                            {
                                indexes.Add(ind - 1);

                                //new used, because we don't want a reference
                                CC_Action action2 = new CC_Action()
                                {
                                    action = actions[ind].action
                                };

                                //new used, because we don't want a reference
                                actions[ind] = new CC_Action()
                                {
                                    action = actions[ind - 1].action
                                };
                                actions[ind - 1] = action2;
                            }
                        }
                    }
                    //when user select from bottom to top
                    else
                    {
                        for (int i = LVactions.SelectedItems.Count - 1; i >= 0; i--)
                        {
                            int ind = LVactions.Items.IndexOf(LVactions.SelectedItems[i]);

                            if (ind != 0)
                            {
                                indexes.Add(ind - 1);

                                //new used, because we don't want a reference
                                CC_Action action2 = new CC_Action()
                                {
                                    action = actions[ind].action
                                };

                                //new used, because we don't want a reference
                                actions[ind] = new CC_Action()
                                {
                                    action = actions[ind - 1].action
                                };
                                actions[ind - 1] = action2;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC008", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cv_actions.Refresh();

                foreach (int ind in indexes)
                    LVactions.SelectedItems.Add(LVactions.Items[ind]);
            }
        }

        private void Bmove_down_Click(object sender, RoutedEventArgs e)
        {
            List<int> indexes = new List<int>();

            try
            {
                if (LVactions.SelectedIndex != -1)
                {
                    int ind_first = LVactions.Items.IndexOf(LVactions.SelectedItems[0]);
                    int ind_last = LVactions.Items.IndexOf(LVactions.SelectedItems[LVactions.SelectedItems.Count - 1]);

                    //when user select from bottom to top
                    if (ind_first >= ind_last)
                    {
                        foreach (CC_Action action in LVactions.SelectedItems)
                        {
                            int ind = LVactions.Items.IndexOf(action);

                            if (ind != LVactions.Items.Count - 1)
                            {
                                indexes.Add(ind + 1);

                                //new used, because we don't want a reference
                                CC_Action action2 = new CC_Action()
                                {
                                    action = actions[ind].action
                                };

                                //new used, because we don't want a reference
                                actions[ind] = new CC_Action()
                                {
                                    action = actions[ind + 1].action
                                };
                                actions[ind + 1] = action2;
                            }
                        }
                    }
                    //when user select from top to bottom
                    else
                    {
                        for (int i = LVactions.SelectedItems.Count - 1; i >= 0; i--)
                        {
                            int ind = LVactions.Items.IndexOf(LVactions.SelectedItems[i]);

                            if (ind != LVactions.Items.Count - 1)
                            {
                                indexes.Add(ind + 1);

                                //new used, because we don't want a reference
                                CC_Action action2 = new CC_Action()
                                {
                                    action = actions[ind].action
                                };

                                //new used, because we don't want a reference
                                actions[ind] = new CC_Action()
                                {
                                    action = actions[ind + 1].action
                                };
                                actions[ind + 1] = action2;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC009", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cv_actions.Refresh();

                foreach (int ind in indexes)
                    LVactions.SelectedItems.Add(LVactions.Items[ind]);
            }
        }

        private void Badd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVnew_actions.SelectedIndex;

                if (ind != -1)
                {
                    string action_text = ((CC_Action)LVnew_actions.Items[ind]).action;

                    if (action_text == "Press key(s)")
                    {
                        WindowAddEditActionKeyboard w = new WindowAddEditActionKeyboard(
                            new ActionKeyboard() { action_text = "", keys = null, option = 0, time = 0 });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Toggle key(s)")
                    {
                        WindowAddEditActionKeyboard w = new WindowAddEditActionKeyboard(
                            new ActionKeyboard() { action_text = "", keys = null, option = 1, time = 0 });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Hold down key(s)")
                    {
                        WindowAddEditActionKeyboard w = new WindowAddEditActionKeyboard(
                            new ActionKeyboard() { action_text = "", keys = null, option = 2, time = 0 });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Release key(s)")
                    {
                        WindowAddEditActionKeyboard w = new WindowAddEditActionKeyboard(
                            new ActionKeyboard() { action_text = "", keys = null, option = 3, time = 0 });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Release all keys and buttons")
                    {
                        int insert_index = LVactions.SelectedIndex + 1;

                        if (LVactions.SelectedIndex != -1)
                            actions.Insert(insert_index, new CC_Action()
                            { action = "Release all keys and buttons" });
                        else
                            actions.Add(new CC_Action() { action = "Release all keys and buttons" });

                        cv_actions.Refresh();

                        if (insert_index != 0)
                            LVactions.SelectedIndex = insert_index;
                        else
                            LVactions.SelectedIndex = LVactions.Items.Count - 1;

                        LVactions.ScrollIntoView(LVactions.SelectedItem);
                    }
                    else if (action_text == "Click mouse button")
                    {
                        WindowAddEditActionMouse w = new WindowAddEditActionMouse(
                            new ActionMouse()
                            {
                                action_text = "",
                                keys = null,
                                left = true,
                                option = 0,
                                time = 0,
                                x = -1,
                                y = -1
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Toggle mouse button")
                    {
                        WindowAddEditActionMouse w = new WindowAddEditActionMouse(
                            new ActionMouse()
                            {
                                action_text = "",
                                keys = null,
                                left = true,
                                option = 1,
                                time = 0,
                                x = -1,
                                y = -1
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Hold down mouse button")
                    {
                        WindowAddEditActionMouse w = new WindowAddEditActionMouse(
                            new ActionMouse()
                            {
                                action_text = "",
                                keys = null,
                                left = true,
                                option = 2,
                                time = 0,
                                x = -1,
                                y = -1
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Release mouse button")
                    {
                        WindowAddEditActionMouse w = new WindowAddEditActionMouse(
                            new ActionMouse()
                            {
                                action_text = "",
                                keys = null,
                                left = true,
                                option = 3,
                                time = 0,
                                x = -1,
                                y = -1
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Move mouse cursor")
                    {
                        WindowAddEditActionMoveMouse w = new WindowAddEditActionMoveMouse(
                            new ActionMoveMouse()
                            {
                                action_text = "",
                                keys = null,
                                absolute = true,
                                x = -1,
                                y = -1
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Scroll mouse wheel")
                    {
                        WindowAddEditActionScrollMouse w = new WindowAddEditActionScrollMouse(
                            new ActionScrollMouse()
                            {
                                action_text = "",
                                keys = null,
                                up = true,
                                scrolling_value = 40
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Open file/program")
                    {
                        Microsoft.Win32.RegistryKey reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                    .OpenSubKey(Middle_Man.registry_path_easy, true);

                        string path = null;

                        if (reg_key_easy != null)
                        {
                            object ob = reg_key_easy.GetValue(Middle_Man.registry_key_last_open_program_path);

                            if (ob != null)
                            {
                                path = ob.ToString();
                            }
                        }

                        OpenFileDialog ofd = new OpenFileDialog();
                        ofd.Title = "Choose a file/program";
                        ofd.Multiselect = true;
                        ofd.Filter = "All files (*.*)|*.*";

                        if (path == null)
                            ofd.InitialDirectory = Environment.GetFolderPath(
                                Environment.SpecialFolder.Desktop);
                        else
                            ofd.InitialDirectory = path;

                        if (ofd.ShowDialog() == true)
                        {
                            if (reg_key_easy == null)
                            {
                                reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                        .CreateSubKey(Middle_Man.registry_path_easy, true);
                            }

                            reg_key_easy.SetValue(Middle_Man.registry_key_last_open_program_path,
                                ofd.FileName);

                            int insert_index = LVactions.SelectedIndex + 1;

                            foreach (string filename in ofd.FileNames)
                            {
                                if (LVactions.SelectedIndex != -1)
                                    actions.Insert(insert_index,
                                        new CC_Action() { action = "Open: " + filename });
                                else
                                    actions.Add(new CC_Action() { action = "Open: " + filename });

                                if (insert_index != 0)
                                    insert_index++;
                            }
                            if (insert_index != 0)
                                insert_index--;

                            cv_actions.Refresh();

                            if (insert_index != 0)
                                LVactions.SelectedIndex = insert_index;
                            else
                                LVactions.SelectedIndex = LVactions.Items.Count - 1;

                            LVactions.ScrollIntoView(LVactions.SelectedItem);
                        }
                    }
                    else if (action_text == "Open URL")
                    {
                        WindowAddEditActionOpenURL w = new WindowAddEditActionOpenURL(
                            new ActionOpenURL()
                            {
                                action_text = "",
                                url = ""
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Play sound")
                    {
                        Microsoft.Win32.RegistryKey reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                    .OpenSubKey(Middle_Man.registry_path_easy, true);

                        string path = null;

                        if (reg_key_easy != null)
                        {
                            object ob = reg_key_easy.GetValue(
                                Middle_Man.registry_key_last_open_sound_file_path);

                            if (ob != null)
                            {
                                path = ob.ToString();
                            }
                        }

                        OpenFileDialog ofd = new OpenFileDialog();
                        ofd.Title = "Choose an mp3 file";
                        ofd.Multiselect = true;
                        ofd.Filter = "MP3 files (*.mp3)|*.mp3|All files (*.*)|*.*";

                        if (path == null)
                            ofd.InitialDirectory = Environment.GetFolderPath(
                                Environment.SpecialFolder.MyMusic);
                        else
                            ofd.InitialDirectory = path;

                        if (ofd.ShowDialog() == true)
                        {
                            if (reg_key_easy == null)
                            {
                                reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                        .CreateSubKey(Middle_Man.registry_path_easy, true);
                            }

                            reg_key_easy.SetValue(Middle_Man.registry_key_last_open_sound_file_path,
                                ofd.FileName);

                            int insert_index = LVactions.SelectedIndex + 1;

                            foreach (string filename in ofd.FileNames)
                            {
                                if (LVactions.SelectedIndex != -1)
                                    actions.Insert(insert_index,
                                        new CC_Action() { action = "Play sound: " + filename });
                                else
                                    actions.Add(new CC_Action() { action = "Play sound: " + filename });

                                if (insert_index != 0)
                                    insert_index++;
                            }
                            if (insert_index != 0)
                                insert_index--;

                            cv_actions.Refresh();

                            if (insert_index != 0)
                                LVactions.SelectedIndex = insert_index;
                            else
                                LVactions.SelectedIndex = LVactions.Items.Count - 1;

                            LVactions.ScrollIntoView(LVactions.SelectedItem);
                        }
                    }
                    else if (action_text == "Read aloud text")
                    {
                        WindowAddEditActionReadText w = new WindowAddEditActionReadText(
                            new ActionReadAloud()
                            {
                                action_text = "",
                                text = ""
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Type text")
                    {
                        WindowAddEditActionTypeText w = new WindowAddEditActionTypeText(
                            new ActionTypeText()
                            {
                                action_text = "",
                                text = ""
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                    else if (action_text == "Wait")
                    {
                        WindowAddEditActionWait w = new WindowAddEditActionWait(
                            new ActionWait()
                            {
                                action_text = "",
                                time = 1000
                            });
                        w.Owner = Application.Current.MainWindow;
                        w.ShowInTaskbar = false;
                        w.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC010", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string name = TBname.Text;
                string description = TBdescription.Text;
                string group;
                string max_executions = TBmax_executions.Text;
                bool enabled = (bool)CHBenabled.IsChecked;
                int ind = Middle_Man.get_profile_ind_by_name(cc_profile);
                short trash = 0;
                int command_ind = -1;

                if (name == "")
                    throw new Exception("Command name cannot be empty.");

                foreach (CustomCommand cc in Middle_Man.profiles[ind].custom_commands)
                {
                    if (cc.name == name && cc.name != prev_name)
                    {
                        throw new Exception("A command with the same name already exists in" +
                            " this profile.");
                    }
                }

                if (CBgroup.SelectedIndex == -1)
                {
                    throw new Exception("No group is selected.");
                }
                else
                {
                    group = CBgroup.SelectedItem.ToString();
                }

                if ((short.TryParse(max_executions, out trash) == false) || trash <= 0 || trash > 50)
                {
                    throw new Exception("Max executions must be a number between 1 and 50.");
                }

                if (TM && edit == false)
                {
                    int n = 0;

                    foreach (Profile p in Middle_Man.profiles)
                    {
                        n += p.custom_commands.Count;
                    }

                    if (n >= CCL)
                    {
                        throw new Exception("The trial version allows having up to " + CCL +
                            " custom commands in all profiles.");
                    }
                }

                List<CC_Action> actions = new List<CC_Action>();

                foreach (CC_Action action in LVactions.Items)
                {
                    actions.Add(action);
                }

                if (edit == false)
                {
                    Middle_Man.profiles[ind].custom_commands.Add(new CustomCommand()
                    {
                        name = name,
                        description = description,
                        group = group,
                        max_executions = short.Parse(max_executions),
                        enabled = enabled,
                        actions = actions
                    });
                    command_ind = Middle_Man.profiles[ind].custom_commands.Count - 1;
                }
                else
                {
                    for (int i = 0; i < Middle_Man.profiles[ind].custom_commands.Count; i++)
                    {
                        if (Middle_Man.profiles[ind].custom_commands[i].name == prev_name)
                        {
                            Middle_Man.profiles[ind].custom_commands[i].name = name;
                            Middle_Man.profiles[ind].custom_commands[i].description = description;
                            Middle_Man.profiles[ind].custom_commands[i].group = group;
                            Middle_Man.profiles[ind].custom_commands[i].max_executions
                                = short.Parse(max_executions);
                            Middle_Man.profiles[ind].custom_commands[i].enabled = enabled;
                            Middle_Man.profiles[ind].custom_commands[i].actions = actions;

                            command_ind = i;

                            break;
                        }
                    }
                }

                Middle_Man.sort_commands_by_name_asc(ind);
                Middle_Man.save_profiles();

                foreach (System.Windows.Window window in Application.Current.Windows)
                {
                    if (window.GetType() == typeof(MainWindow))
                    {
                        MainWindow w = (MainWindow)window;

                        w.cv_LVcommands.Refresh();

                        //sorting may changed index
                        command_ind = Middle_Man.get_command_ind_by_name(ind, name);

                        w.LVcommands.SelectedIndex = command_ind;

                        w.LVcommands.ScrollIntoView(w.LVcommands.SelectedItem);

                        w.LVactions.ItemsSource = Middle_Man.profiles[ind].custom_commands[command_ind].actions;

                        CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(
                                    w.LVactions.ItemsSource);
                        cv.Refresh();

                        if (edit)
                            w.LVcommands.Focus();
                        else
                            w.Badd_command.Focus();
                    }
                }

                Middle_Man.last_used_max_executions = max_executions;

                Middle_Man.force_updating_both_cc_grammars = true;

                this.Close();
            }
            catch (Exception ex)
            {
                if (ex.Message.StartsWith("The trial version allows"))
                    MessageBox.Show(ex.Message, "Trial version limit reached",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                else
                    MessageBox.Show(ex.Message, "Error WAEC011", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bcancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC012", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TBdescription_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            try
            {
                if (TBdescription.Text.Length == 0 && Loptional.Visibility == Visibility.Hidden)
                    Loptional.Visibility = Visibility.Visible;
                else if (TBdescription.Text.Length > 0 && Loptional.Visibility == Visibility.Visible)
                    Loptional.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC013", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Delete)
                {
                    Bdelete_Click(null, null);
                }
                else if (e.Key == Key.C
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIcopy_actions_Click(null, null);
                }
                else if (e.Key == Key.V
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    MIpaste_actions_Click(null, null);
                }
                else if (e.Key == Key.Up
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control
                    && Bmove_up.IsEnabled)
                {
                    MImove_up_actions_Click(null, null);
                }
                else if (e.Key == Key.Down
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control
                    && Bmove_down.IsEnabled)
                {
                    MImove_down_actions_Click(null, null);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC0014", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVactions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                int ind = LVactions.SelectedIndex;

                if (ind == -1)
                {
                    Bedit.IsEnabled = false;
                    Bcopy.IsEnabled = false;
                    Bdelete.IsEnabled = false;
                    Bmove_up.IsEnabled = false;
                    Bmove_down.IsEnabled = false;

                    MIedit_actions.IsEnabled = false;
                    MIcopy_actions.IsEnabled = false;
                    MIpaste_actions.IsEnabled = false;
                    MIdelete_actions.IsEnabled = false;
                    MImove_up_actions.IsEnabled = false;
                    MImove_down_actions.IsEnabled = false;
                }
                else
                {
                    if (((CC_Action)LVactions.SelectedItem).action.StartsWith("Release all keys") == false)
                    {
                        Bedit.IsEnabled = true;
                        MIedit_actions.IsEnabled = true;
                    }

                    Bcopy.IsEnabled = true;
                    Bdelete.IsEnabled = true;

                    MIcopy_actions.IsEnabled = true;
                    MIdelete_actions.IsEnabled = true;

                    bool first_selected = false; //in multi selection
                    bool last_selected = false;  //in multi selection

                    int ind_first = LVactions.Items.IndexOf(LVactions.SelectedItems[0]);
                    int ind_last = LVactions.Items.IndexOf(LVactions.SelectedItems[LVactions.SelectedItems.Count - 1]);

                    //user may select in both directions
                    if (ind_first == 0 || ind_last == 0)
                        first_selected = true;
                    else if (ind_first == LVactions.Items.Count - 1 || ind_last == LVactions.Items.Count - 1)
                        last_selected = true;

                    if (LVactions.Items.Count > 1)
                    {
                        if (first_selected)
                        {
                            Bmove_up.IsEnabled = false;
                            MImove_up_actions.IsEnabled = false;
                        }
                        else
                        {
                            Bmove_up.IsEnabled = true;
                            MImove_up_actions.IsEnabled = true;
                        }

                        if (last_selected)
                        {
                            Bmove_down.IsEnabled = false;
                            MImove_down_actions.IsEnabled = false;
                        }
                        else
                        {
                            Bmove_down.IsEnabled = true;
                            MImove_down_actions.IsEnabled = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC015", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVnew_actions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                int ind = LVnew_actions.SelectedIndex;

                if (ind == -1)
                {
                    Badd.IsEnabled = false;
                }
                else
                {
                    Badd.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC016", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIedit_actions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Bedit_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC017", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIcopy_actions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVactions.SelectedIndex != -1)
                    Bcopy_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC018", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIpaste_actions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Bpaste_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC019", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIdelete_actions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVactions.SelectedIndex != -1)
                    Bdelete_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC020", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MImove_down_actions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVactions.SelectedIndex != -1)
                    Bmove_down_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC021", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MImove_up_actions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVactions.SelectedIndex != -1)
                    Bmove_up_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC022", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVactions_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (LVactions.SelectedIndex != -1)
                    Bedit_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC023", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVnew_actions_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (LVnew_actions.SelectedIndex != -1)
                    Badd_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC024", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Iquestion_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                string text = "To increase the command’s execution number say twice or x times after saying" +
                    " the command name.\r\n" +
                    "For example:\r\n" +
                    "- Say \"My Command twice\" to execute command named \"My Command\" twice.\r\n" +
                    "- Say \"My Command 8 times\" to execute command named \"My Command\" 8 times.\r\n" +
                    "Do not set maximum executions higher than necessary.";

                MessageBox.Show(text, "Information", MessageBoxButton.OK,
                MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC025", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Brecord_actions_Click(object sender, RoutedEventArgs e)
        {
            WindowRecordActions w = new WindowRecordActions();

            try
            {
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC026", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                foreach (System.Windows.Window window in Application.Current.Windows)
                {
                    if (window.GetType() == typeof(WindowRecordActions))
                    {
                        WindowRecordActions w = (WindowRecordActions)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditActionKeyboard))
                    {
                        WindowAddEditActionKeyboard w = (WindowAddEditActionKeyboard)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditActionMouse))
                    {
                        WindowAddEditActionMouse w = (WindowAddEditActionMouse)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditActionMoveMouse))
                    {
                        WindowAddEditActionMoveMouse w = (WindowAddEditActionMoveMouse)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditActionOpenURL))
                    {
                        WindowAddEditActionOpenURL w = (WindowAddEditActionOpenURL)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditActionReadText))
                    {
                        WindowAddEditActionReadText w = (WindowAddEditActionReadText)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditActionScrollMouse))
                    {
                        WindowAddEditActionScrollMouse w = (WindowAddEditActionScrollMouse)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditActionTypeText))
                    {
                        WindowAddEditActionTypeText w = (WindowAddEditActionTypeText)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditActionWait))
                    {
                        WindowAddEditActionWait w = (WindowAddEditActionWait)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowAddEditGroup))
                    {
                        WindowAddEditGroup w = (WindowAddEditGroup)window;

                        w.Close();
                    }
                    if (window.GetType() == typeof(WindowManageGroups))
                    {
                        WindowManageGroups w = (WindowManageGroups)window;

                        w.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEC027", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}