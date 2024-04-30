//highest error nr: AM028
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml;

namespace Speech
{
    public partial class MainWindow : Window
    {
        [DllImport("shell32.dll")]
        static extern bool SHGetSpecialFolderPath(IntPtr hwndOwner,
            [Out] StringBuilder lpszPath, int nFolder, bool fCreate);
        const int CSIDL_COMMON_STARTMENU = 0x16;  // All Users\Start Menu
        //const int CSIDL_PROGRAMS = 2; // not enough programs

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName,
            string lpWindowName);

        //Maksymalizowanie i normalizowanie
        [DllImport("USER32.DLL")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_MAXIMIZE = 3;
        private const int SW_SHOWNORMAL = 1;
        private const int SW_MINIMIZE = 6;
        private const int SW_RESTORE = 9;

        //Not used right now, but could be used in the future:
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const UInt32 SWP_NOSIZE = 0x0001;
        const UInt32 SWP_NOMOVE = 0x0002;
        const UInt32 SWP_SHOWWINDOW = 0x0040;

        System.Windows.Forms.NotifyIcon ni = new System.Windows.Forms.NotifyIcon();

        System.Windows.Forms.ColorDialog colorDialog1 = new System.Windows.Forms.ColorDialog();
        System.Windows.Forms.ColorDialog colorDialog2 = new System.Windows.Forms.ColorDialog();

        public int GetWindowsScaling()
        {
            return (int)(100 * System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width
                / SystemParameters.PrimaryScreenWidth);
        }

        void toggle_grammar(bool enabled, grammar_type gt)
        {
            string grammar_name = "";

            if (gt == grammar_type.grammar_apps_opening)
                grammar_name = grammar_apps_opening_name;
            else if (gt == grammar_type.grammar_apps_switching)
                grammar_name = grammar_apps_switching_name;
            else if (gt == grammar_type.grammar_builtin_commands)
                grammar_name = grammar_builtin_commands_name;
            else if (gt == grammar_type.grammar_custom_commands_any)
                grammar_name = grammar_custom_commands_any_name;
            else if (gt == grammar_type.grammar_custom_commands_foreground)
                grammar_name = grammar_custom_commands_foreground_name;
            else if (gt == grammar_type.grammar_dictation)
                grammar_name = grammar_dictation_name;
            else if (gt == grammar_type.grammar_dictation_commands)
                grammar_name = grammar_dictation_commands_name;
            else if (gt == grammar_type.grammar_mousegrid)
                grammar_name = grammar_mousegrid_name;
            else if (gt == grammar_type.grammar_off_mode)
                grammar_name = grammar_off_mode_name;

            for (int i = 0; i < recognizer.Grammars.Count; i++)
            {
                if (recognizer.Grammars[i].Name == grammar_name)
                {
                    recognizer.Grammars[i].Enabled = enabled;
                    break;
                }
            }
        }

        bool is_grammar_loaded(string grammar_name)
        {
            for (int i = 0; i < recognizer.Grammars.Count; i++)
            {
                if (recognizer.Grammars[i].Name == grammar_name)
                {
                    return true;
                }
            }

            return false;
        }

        void unload_grammar_if_loaded(Grammar grammar, string grammar_name)
        {
            if (is_grammar_loaded(grammar_name))
            {
                recognizer.UnloadGrammar(grammar);
            }
        }

        void restore_default_settings()
        {
            TBconfidence_start.Text = default_confidence_turning_on.ToString();
            TBconfidence_commands.Text = default_confidence_other_commands.ToString();
            TBconfidence_dictation.Text = default_confidence_dictation.ToString();
            CHBuse_better_dictation.IsChecked = default_better_dictation;

            bool found = false;
            for (int i = 0; i < ss_voices_priority_list.Count && found == false; i++)
            {
                foreach (InstalledVoice iv in installed_voices)
                {
                    if (iv.VoiceInfo.Name.Contains(ss_voices_priority_list[i]))
                    {
                        default_ss_voice = iv.VoiceInfo.Name;
                        found = true;
                        break;
                    }
                }
            }

            if (CBss_voices.Items.Count > 0)
            {
                if (default_ss_voice == "")
                {
                    CBss_voices.SelectedIndex = 0;
                    default_ss_voice = CBss_voices.SelectedItem.ToString();
                }
                else
                {
                    CBss_voices.SelectedItem = default_ss_voice;
                }
            }
            ss_voice = default_ss_voice;
            
            TBss_volume.Text = default_ss_volume.ToString();
            CHBread_recognized_speech.IsChecked = default_read_recognized_speech;

            CHBstart_with_hidden.IsChecked = default_start_with_hidden;
            CHBrun_at_startup.IsChecked = default_run_at_startup;
            CHBstart_minimized.IsChecked = default_start_minimized;
            CHBminimize_to_tray.IsChecked = default_minimize_to_tray;
            CHBauto_updates.IsChecked = default_auto_updates;

            CBtype.SelectedItem = default_grid_type.ToString().Replace("_", " ").FirstCharToUpper();
            CBlines.SelectedIndex = default_grid_lines;
            TBdesired_figures_nr.Text = default_desired_figures_nr.ToString();
            color_bg_str = default_color_bg_str;
            color_font_str = default_color_font_str;

            int argb = Convert.ToInt32(color_bg_str);

            byte[] values = BitConverter.GetBytes(argb);

            byte a = values[3];
            byte r = values[2];
            byte g = values[1];
            byte b = values[0];

            TBbackground_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));

            argb = Convert.ToInt32(color_font_str);

            values = BitConverter.GetBytes(argb);

            a = values[3];
            r = values[2];
            g = values[1];
            b = values[0];

            TBfont_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));

            TBfont_size.Text = default_font_size.ToString();
            CHBsmart_mousegrid.IsChecked = default_smart_grid;
        }

        void toggle_better_dictation()
        {
            key_down(WindowsInput.Native.VirtualKeyCode.LWIN);
            key_down(WindowsInput.Native.VirtualKeyCode.VK_H);
            Thread.Sleep(75);
            key_up(WindowsInput.Native.VirtualKeyCode.VK_H);
            key_up(WindowsInput.Native.VirtualKeyCode.LWIN);
        }

        void restart_better_dictation()
        {
            toggle_better_dictation();
            Thread.Sleep(100);
            toggle_better_dictation();
        }

        void CenterWindowOnScreen()
        {
            double screenWidth = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenWidth);
            double screenHeight = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenHeight);
            double windowWidth = this.Width;
            double windowHeight = this.Height;
            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);
        }

        private void PositionSpeechWindow(int x = -1, int y = -1)
        {
            double screenWidth = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenWidth);
            double screenHeight = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenHeight);
            double windowWidth = SW.Width;
            double windowHeight = SW.Height;
            if (x == -1 && y == -1)
            {
                SW.Left = (screenWidth / 2) - (windowWidth / 2);
                SW.Top = (screenHeight) - (windowHeight + 34);
            }
            else
            {
                SW.Left = x;
                SW.Top = y;
            }
        }

        bool dont_compare = false;
        private void Brestore_default_settings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                dont_compare = true;

                restore_default_settings();

                PositionSpeechWindow();

                Bsave_settings_Click(null, null);
                dont_compare = false;

                compare_current_and_default();
                compare_current_and_saved();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void ni_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            try
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Left)
                {
                    Show();
                    this.WindowState = WindowState.Normal;
                    SetForegroundWindow(new System.Windows.Interop.WindowInteropHelper(this).Handle);
                }
                else if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    this.Activate();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void mi_switch_to_command_mode_Click(object sender, EventArgs e)
        {
            try
            {
                if (current_mode == mode.dictation && better_dictation)
                {
                    toggle_better_dictation();
                }

                current_mode = mode.command;
                load_turned_on();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM025", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void mi_switch_to_dictation_mode_Click(object sender, EventArgs e)
        {
            try
            {
                current_mode = mode.dictation;
                load_turned_on();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM026", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void mi_switch_to_off_mode_Click(object sender, EventArgs e)
        {
            try
            {
                if (current_mode == mode.dictation && better_dictation)
                {
                    toggle_better_dictation();
                }

                load_turned_off();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM027", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void mi_exit_Click(object sender, EventArgs e)
        {
            try
            {
                Window_Closing(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM028", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                ni.Visible = false;
                ni.Dispose();

                if (THRmonitor != null && current_mode == mode.dictation && better_dictation)
                {
                    toggle_better_dictation();
                }

                release_buttons_and_keys();

                if (smart_grid)
                    save_grids();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM003", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            Process.GetCurrentProcess().Kill();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                //if (test_mode)
                //{
                //    MessageBox.Show("Work by Speech is running in test mode.", "Warning",
                //        MessageBoxButton.OK, MessageBoxImage.Warning);
                //}

                release_buttons_and_keys();

                if (first_run)
                {
                    MessageBox.Show("This program can recognize your speech with high accuracy only if you " +
                        "complete at least two voice trainings. One voice training takes about 7 minutes. " +
                        "You can find more information about voice training in point 4 of the user guide, " +
                        "which is located in the help section.\n\n"
                        + "To learn how to use this program, go to the help section and read " +
                        "all the documents.",
                        "Welcome",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }

                update_app_if_necessary();

                LVprofiles.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void update_app_if_necessary()
        {
            bool update_available = false;

            if (latest_version != "unknown" &&
                int.Parse(latest_version.Replace(".", "")) > int.Parse(prog_version.Replace(".", "")))
            {
                update_available = true;
            }

            if ((bool)CHBauto_updates.IsChecked && update_available)
            {
                MessageBoxResult dialogResult = System.Windows.MessageBox.Show("A new program version" +
                    " is available. Do you want to download it now?",
                //    " is available. Do you want to perform an automatic update now?",
                    "New Version Available", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (dialogResult == MessageBoxResult.Yes)
                {
                    //Automatic updating (removed thanks to security vendors which don't accept
                    //auto updaters that don't have EV certificate like it's for free)
                    /*
                    Process.Start("Updater.exe");
                    Process.GetCurrentProcess().Kill();
                    */

                    //Open download page
                    Process.Start(Middle_Man.url_download);
                }
            }
        }

        void TBbackground_color_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                System.Windows.Forms.DialogResult dr = colorDialog1.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    int argb = Convert.ToInt32(colorDialog1.Color.ToArgb().ToString());

                    byte[] values = BitConverter.GetBytes(argb);

                    byte a = values[3];
                    byte r = values[2];
                    byte g = values[1];
                    byte b = values[0];

                    TBbackground_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));

                    update_mousegrid_preview();

                    compare_current_and_default();
                    compare_current_and_saved();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TBfont_color_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                System.Windows.Forms.DialogResult dr = colorDialog2.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    int argb = Convert.ToInt32(colorDialog2.Color.ToArgb().ToString());

                    byte[] values = BitConverter.GetBytes(argb);

                    byte a = values[3];
                    byte r = values[2];
                    byte g = values[1];
                    byte b = values[0];

                    TBfont_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));

                    update_mousegrid_preview();

                    compare_current_and_default();
                    compare_current_and_saved();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TBfont_size_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (TBmousegrid_preview != null)
            {
                if (int.TryParse(TBfont_size.Text, out int result))
                {
                    if (result > 0 && result <= max_font_size)
                    {
                        update_mousegrid_preview(result);

                        compare_current_and_default();
                        compare_current_and_saved();
                    }
                }
            }
        }

        private void CBtype_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CBtype.SelectedItem.ToString() == "Hexagonal")
            {
                CBlines.SelectedIndex = 0;
                CBlines.IsEnabled = false;
            }
            else
            {
                CBlines.IsEnabled = true;
            }

            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBread_recognized_speech_Checked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBuse_better_dictation_Checked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBsmart_mousegrid_Checked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBstart_with_hidden_Checked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBrun_at_startup_Checked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBstart_minimized_Checked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBminimize_to_tray_Checked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBauto_updates_Checked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBuse_better_dictation_Unchecked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBread_recognized_speech_Unchecked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBstart_with_hidden_Unchecked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBrun_at_startup_Unchecked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBstart_minimized_Unchecked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBminimize_to_tray_Unchecked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBauto_updates_Unchecked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CHBsmart_mousegrid_Unchecked(object sender, RoutedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void TBconfidence_start_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void TBconfidence_commands_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void TBconfidence_dictation_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void TBdesired_figures_nr_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CBss_voices_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void TBss_volume_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void CBlines_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            compare_current_and_default();
            compare_current_and_saved();
        }

        private void W_StateChanged(object sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Minimized && CHBminimize_to_tray.IsChecked == true)
            {
                this.Hide();
            }
        }

        private void Bsave_settings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                while (recognition_suspended)
                {
                    Thread.Sleep(10);
                }

                bool prev_check_for_updates_status = auto_updates;

                recognition_suspended = true;
                
                Mouse.OverrideCursor = Cursors.Wait;
                SW.Bmode.Visibility = Visibility.Hidden;

                save_settings();

                recognition_suspended = false;
                
                SW.Bmode.Visibility = Visibility.Visible;
                Mouse.OverrideCursor = null;

                if (dont_compare == false)
                {
                    compare_current_and_default();
                    compare_current_and_saved();
                }

                if (prev_check_for_updates_status == false && (bool)CHBauto_updates.IsChecked)
                {
                    update_app_if_necessary();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM007", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bchangelog_Click(object sender, RoutedEventArgs e)
        {
            WindowChangelog w = new WindowChangelog();
            w.Owner = Application.Current.MainWindow;
            w.ShowInTaskbar = false;
            w.ShowDialog();
        }

        private void Beula_Click(object sender, RoutedEventArgs e)
        {
            WindowEULA w = new WindowEULA();
            w.Owner = Application.Current.MainWindow;
            w.ShowInTaskbar = false;
            w.ShowDialog();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("User Guide.pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM009", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("Built-in Commands.pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM010", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("Mousegrid Alphabet.pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM011", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("Useful Windows Key Combinations.pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM012", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Lhomepage_PreviewMouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                Process.Start(Middle_Man.url_homepage_full);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM013", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Lhomepage_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
        }

        private void Lhomepage_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Mouse.OverrideCursor = null;
        }

        
        private void Iquestion2_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                MessageBox.Show(Iquestion2.ToolTip.ToString(), "Information", MessageBoxButton.OK,
                MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM024", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void update_mousegrid_preview(int font_size_temp = 0)
        {
            if (font_size_temp == 0)
                font_size_temp = font_size;

            TBmousegrid_preview.Background = TBbackground_color.Background;
            TBmousegrid_preview.Foreground = TBfont_color.Background;
            TBmousegrid_preview.FontSize = font_size_temp;

            Size size = MeasureString(TBmousegrid_preview.Text, TBmousegrid_preview);
            TBmousegrid_preview.Width = size.Width + 2;
            TBmousegrid_preview.Height = size.Height + 2;
        }

        void set_values()
        {
            if (saving_enabled)
            {
                if (smart_grid && desired_figures_nr >= 5)
                    save_grids();
            }

            bool grid_size_changed = false;
            bool prev_better_dictation = better_dictation;
            GridType prev_grid_type = grid_type;

            confidence_turning_on = int.Parse(TBconfidence_start.Text);
            confidence_other_commands = int.Parse(TBconfidence_commands.Text);
            confidence_dictation = int.Parse(TBconfidence_dictation.Text);
            better_dictation = (bool)CHBuse_better_dictation.IsChecked;

            if(CBss_voices.Items.Count > 0)
                ss_voice = CBss_voices.SelectedItem.ToString();
            if (ss_voice != "")
                ss.SelectVoice(ss_voice);

            ss_volume = int.Parse(TBss_volume.Text);
            read_recognized_speech = (bool)CHBread_recognized_speech.IsChecked;
            
            start_with_hidden = (bool)CHBstart_with_hidden.IsChecked;
            run_at_startup = (bool)CHBrun_at_startup.IsChecked;
            start_minimized = (bool)CHBstart_minimized.IsChecked;
            minimize_to_tray = (bool)CHBminimize_to_tray.IsChecked;
            auto_updates = (bool)CHBauto_updates.IsChecked;

            if (CBtype.SelectedIndex == 0)
                grid_type = GridType.hexagonal;
            else if (CBtype.SelectedIndex == 1)
                grid_type = GridType.square;
            else if (CBtype.SelectedIndex == 2)
                grid_type = GridType.square_horizontal_precision;
            else if (CBtype.SelectedIndex == 3)
                grid_type = GridType.square_vertical_precision;
            else if (CBtype.SelectedIndex == 4)
                grid_type = GridType.square_combined_precision;

            grid_lines = CBlines.SelectedIndex;

            //if grid size changed whole grid must be generated again
            if (desired_figures_nr != int.Parse(TBdesired_figures_nr.Text)
                || grid_type != prev_grid_type)
            {
                grid_size_changed = true;
            }

            desired_figures_nr = int.Parse(TBdesired_figures_nr.Text);

            if (saving_enabled == false)
            {
                int argb = Convert.ToInt32(color_bg_str);

                byte[] values = BitConverter.GetBytes(argb);

                byte a = values[3];
                byte r = values[2];
                byte g = values[1];
                byte b = values[0];

                TBbackground_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));
                color_bg = Color.FromArgb(a, r, g, b);

                argb = Convert.ToInt32(color_font_str);

                values = BitConverter.GetBytes(argb);

                a = values[3];
                r = values[2];
                g = values[1];
                b = values[0];

                TBfont_color.Background = new SolidColorBrush(Color.FromArgb(a, r, g, b));
                color_font = Color.FromArgb(a, r, g, b);
            }
            else
            {
                SolidColorBrush scb = (SolidColorBrush)TBbackground_color.Background;
                System.Drawing.Color color =
                    System.Drawing.Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
                color_bg_str = color.ToArgb().ToString();
                color_bg = Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);

                scb = (SolidColorBrush)TBfont_color.Background;
                color =
                    System.Drawing.Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
                color_font_str = color.ToArgb().ToString();
                color_font = Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
            }

            font_size = int.Parse(TBfont_size.Text);

            update_mousegrid_preview();

            bool prev_smart_grid = smart_grid;
            smart_grid = (bool)CHBsmart_mousegrid.IsChecked;

            //need a .bat file to start an .exe file for some reasons
            Microsoft.Win32.RegistryKey rkApp = Microsoft.Win32.Registry.CurrentUser
                .OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            if (CHBrun_at_startup.IsChecked == true)
            {
                if (rkApp.GetValue(Middle_Man.prog_name) == null)
                {
                    rkApp.SetValue(Middle_Man.prog_name,
                        System.Reflection.Assembly.GetExecutingAssembly().Location.
                        Replace(".exe", ".vbs"));
                }
                generate_bat_file();
            }
            else if (rkApp.GetValue(Middle_Man.prog_name) != null)
            {
                rkApp.DeleteValue(Middle_Man.prog_name, false);
            }

            if ((grid_size_changed || prev_smart_grid != smart_grid)
                && THRswitch_to != null && THRmonitor != null) //we don't want this executed when settings are being loaded
            {
                if (dont_unload_grammars)
                {
                    restart_recognizer(true);
                }
                else
                {
                    //recognizer.RequestRecognizerUpdate();
                    
                    unload_grammar_if_loaded(grammar_mousegrid, grammar_mousegrid_name);

                    grammar_mousegrid = create_grid_grammar(true);

                    if (grammar_mousegrid != null)
                        recognizer.LoadGrammar(grammar_mousegrid);

                    //loaded grammar is enabled by default
                    if(current_mode == mode.grid)
                        toggle_grammar(true, grammar_type.grammar_mousegrid);
                    else
                        toggle_grammar(false, grammar_type.grammar_mousegrid);
                }

                if ((smart_grid && prev_smart_grid == false) || grid_size_changed)
                {
                    load_grids();
                }
            }

            //better dictation uses different grammar
            if (prev_better_dictation != better_dictation
                && THRswitch_to != null && THRmonitor != null) //we don't want this executed when settings are being loaded
            {
                if (dont_unload_grammars)
                {
                    restart_recognizer(false);
                }
                else
                {
                    //recognizer.RequestRecognizerUpdate();
                    
                    unload_grammar_if_loaded(grammar_dictation_commands, grammar_dictation_commands_name);

                    if (are_all_bic_dictation_disabled() == false)
                        grammar_dictation_commands = create_dictation_commands_grammar();
                    else
                        grammar_dictation_commands = null;

                    if (grammar_dictation_commands != null)
                        recognizer.LoadGrammar(grammar_dictation_commands);

                    if(current_mode == mode.dictation)
                    {
                        toggle_grammar(true, grammar_type.grammar_dictation_commands);
                        
                        toggle_better_dictation();
                    }
                    else
                    {
                        toggle_grammar(false, grammar_type.grammar_dictation_commands);
                    }
                }
            }

            //only needed if grid was already created //we don't want this executed when settings
            //are being loaded
            if (grid_width != 0)
            {
                //bad idea:
                //if (auto_grid_font_size)
                //{
                //    if (grid_type == GridType.hexagonal)
                //        font_size = (int)(Math.Floor(figure_width * 14 / 35.555555555555557));
                //    else
                //        font_size = (int)(Math.Floor(figure_width * 14 / 25.098039215686274));
                //}
                if (MW != null)
                {
                    //need to close mousegrid window or there will be 1 more window every time you change
                    //mousegrid settings that require mousegrid regeneration
                    //you can easily see this windows by pressing  windows key + tab
                    MW.Close();
                }
                MW = new MouseGrid(grid_width, grid_height, grid_lines, grid_type, font_family,
                    font_size, color_bg, color_font, rows_nr, cols_nr, figure_width, figure_height,
                    grids[0].elements);
            }
        }

        void generate_bat_file()
        {
            FileStream fs = null;
            StreamWriter sw = null;
            string file_path = System.IO.Path.Combine(
                System.Reflection.Assembly.GetExecutingAssembly().Location.
                        Replace(".exe", ".vbs"));

            try
            {
                fs = new FileStream(file_path, FileMode.Create, FileAccess.Write);
                sw = new StreamWriter(fs);

                //cmd script (black window appearing for a moment is a problem)
                //sw.WriteLine("cd \"" + app_folder_path + "\"");
                //sw.WriteLine("start " 
                //    + System.Diagnostics.Process.GetCurrentProcess().MainModule.ModuleName);

                //vbs script (no window appearing during execution)
                sw.WriteLine("WScript.Sleep 1000");
                sw.WriteLine("Set objShell = CreateObject(\"Wscript.Shell\")");
                sw.WriteLine("objShell.CurrentDirectory = \"" + Middle_Man.app_folder_path + "\"");
                sw.WriteLine("strApp = \"\"\""
                    + System.Diagnostics.Process.GetCurrentProcess().MainModule.ModuleName + "\"\"\"");
                sw.WriteLine("objShell.Run(strApp)");

                sw.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM014", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    if (sw != null)
                        sw.Close();
                    if (fs != null)
                        fs.Close();
                }
                catch (Exception ex2) { }
            }
        }

        void save_settings()
        {
            if (saving_enabled)
            {
                string file_path = System.IO.Path.Combine(Middle_Man.saving_folder_path, filename_settings);

                try
                {
                    int trash;

                    if (int.TryParse(TBconfidence_start.Text, out trash) == false ||
                        int.Parse(TBconfidence_start.Text) < 0 || int.Parse(TBconfidence_start.Text) > 99)
                        throw new Exception("Confidence required for \"Start speech recognition\"" +
                            " command must be between 0 and 99");

                    if (int.TryParse(TBconfidence_commands.Text, out trash) == false 
                        || int.Parse(TBconfidence_commands.Text) < 0
                        || int.Parse(TBconfidence_commands.Text) > 99)
                        throw new Exception("Confidence required for other commands" +
                            " must be between 0 and 99");

                    if (int.TryParse(TBconfidence_dictation.Text, out trash) == false 
                        || int.Parse(TBconfidence_dictation.Text) < 0
                        || int.Parse(TBconfidence_dictation.Text) > 99)
                        throw new Exception("Confidence required for dictation" +
                            " must be between 0 and 99");

                    if (CBss_voices.SelectedIndex == -1 && (bool)CHBread_recognized_speech.IsChecked)
                        throw new Exception("Speech synthesis voice must be selected when" +
                            " \"Read recognized speech\" is checked.");

                    if (int.TryParse(TBss_volume.Text, out trash) == false
                        || int.Parse(TBss_volume.Text) < 0
                        || int.Parse(TBss_volume.Text) > 100)
                        throw new Exception("Speech synthesis volume" +
                            " must be between 0 and 100");

                    if (CBtype.SelectedIndex == -1)
                        throw new Exception("Mousegrid type was not selected.");

                    if (CBlines.SelectedIndex == -1)
                        throw new Exception("Mousegrid lines were not selected.");

                    if (int.TryParse(TBdesired_figures_nr.Text, out trash) == false 
                        || int.Parse(TBdesired_figures_nr.Text) < 5
                        || int.Parse(TBdesired_figures_nr.Text) > max_figures_nr)
                        throw new Exception("Desired figures number must be between 5 and "
                            + max_figures_nr + ".");

                    if (int.TryParse(TBfont_size.Text, out trash) == false 
                        || int.Parse(TBfont_size.Text) < 1
                        || int.Parse(TBfont_size.Text) > max_font_size)
                        throw new Exception("Mousegrid font size must be between 1 and "
                            + max_font_size + ".");

                    if (Directory.Exists(Middle_Man.saving_folder_path) == false)
                    {
                        Directory.CreateDirectory(Middle_Man.saving_folder_path);
                    }

                    set_values();

                    XmlDocument xml_doc = new XmlDocument();

                    XmlNode root_node = xml_doc.CreateElement("settings");

                    XmlAttribute attribute = xml_doc.CreateAttribute("version");
                    attribute.Value = "1";
                    root_node.Attributes.Append(attribute);

                    xml_doc.AppendChild(root_node);

                    XmlNode setting_node;

                    setting_node = xml_doc.CreateElement("confidence_start");
                    setting_node.InnerText = TBconfidence_start.Text;
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("confidence_commands");
                    setting_node.InnerText = TBconfidence_commands.Text;
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("confidence_dictation");
                    setting_node.InnerText = TBconfidence_dictation.Text;
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("read_aloud_recognized_speech");
                    setting_node.InnerText = CHBread_recognized_speech.IsChecked.ToString();
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("use_better_dictation");
                    setting_node.InnerText = CHBuse_better_dictation.IsChecked.ToString();
                    root_node.AppendChild(setting_node);

                    setting_node = xml_doc.CreateElement("type");
                    setting_node.InnerText = CBtype.SelectedIndex.ToString();
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("lines");
                    setting_node.InnerText = CBlines.SelectedIndex.ToString();
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("desired_figures_nr");
                    setting_node.InnerText = TBdesired_figures_nr.Text;
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("color_bg");
                    setting_node.InnerText = color_bg_str;
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("color_font");
                    setting_node.InnerText = color_font_str;
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("font_size");
                    setting_node.InnerText = TBfont_size.Text;
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("smart_mousegrid");
                    setting_node.InnerText = CHBsmart_mousegrid.IsChecked.ToString();
                    root_node.AppendChild(setting_node);

                    setting_node = xml_doc.CreateElement("start_with_hidden");
                    setting_node.InnerText = CHBstart_with_hidden.IsChecked.ToString();
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("run_at_startup");
                    setting_node.InnerText = CHBrun_at_startup.IsChecked.ToString();
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("start_minimized");
                    setting_node.InnerText = CHBstart_minimized.IsChecked.ToString();
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("minimize_to_tray");
                    setting_node.InnerText = CHBminimize_to_tray.IsChecked.ToString();
                    root_node.AppendChild(setting_node);
                    setting_node = xml_doc.CreateElement("auto_updates");
                    setting_node.InnerText = CHBauto_updates.IsChecked.ToString();
                    root_node.AppendChild(setting_node);

                    //added by version 1.3
                    setting_node = xml_doc.CreateElement("ss_voices");
                    if (CBss_voices.Items.Count > 0)
                        setting_node.InnerText = CBss_voices.SelectedItem.ToString();
                    else
                        setting_node.InnerText = "null";
                    root_node.AppendChild(setting_node);

                    //added by version 1.7
                    setting_node = xml_doc.CreateElement("ss_volume");
                    setting_node.InnerText = TBss_volume.Text;
                    root_node.AppendChild(setting_node);

                    xml_doc.Save(Path.Combine(Middle_Man.saving_folder_path, filename_settings));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error AM015", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        void load_settings()
        {
            string file_path = System.IO.Path.Combine(Middle_Man.saving_folder_path, filename_settings);
            string old_file_path = System.IO.Path.Combine(Middle_Man.saving_folder_path, old_filename_settings);

            try
            {
                if (File.Exists(file_path))
                {
                    XmlDocument xml_doc = new XmlDocument();
                    xml_doc.Load(Path.Combine(Middle_Man.saving_folder_path, filename_settings));

                    XmlNodeList settings = xml_doc.SelectNodes("//settings");

                    int version = -1;
                    bool parsing_v = false;

                    if (settings[0].Attributes["version"] != null)
                        parsing_v = int.TryParse(settings[0].Attributes["version"].Value, out version);

                    if (parsing_v && version == 1)
                    {
                        XmlNodeList nodes = settings[0].ChildNodes;

                        foreach (XmlNode node in nodes)
                        {
                            if (node.Name == "confidence_start")
                                TBconfidence_start.Text = node.InnerText;
                            else if (node.Name == "confidence_commands")
                                TBconfidence_commands.Text = node.InnerText;
                            else if (node.Name == "confidence_dictation")
                                TBconfidence_dictation.Text = node.InnerText;
                            else if (node.Name == "read_aloud_recognized_speech")
                                CHBread_recognized_speech.IsChecked = bool.Parse(node.InnerText);
                            else if (node.Name == "use_better_dictation")
                                CHBuse_better_dictation.IsChecked = better_dictation = 
                                    bool.Parse(node.InnerText);
                            else if (node.Name == "type")
                                CBtype.SelectedIndex = int.Parse(node.InnerText);
                            else if (node.Name == "lines")
                                CBlines.SelectedIndex = int.Parse(node.InnerText);
                            else if (node.Name == "desired_figures_nr")
                                TBdesired_figures_nr.Text = node.InnerText;
                            else if (node.Name == "color_bg")
                                color_bg_str = node.InnerText;
                            else if (node.Name == "color_font")
                                color_font_str = node.InnerText;
                            else if (node.Name == "font_size")
                                TBfont_size.Text = node.InnerText;
                            else if (node.Name == "smart_mousegrid")
                                CHBsmart_mousegrid.IsChecked = smart_grid = bool.Parse(node.InnerText);
                            else if (node.Name == "start_with_hidden")
                                CHBstart_with_hidden.IsChecked = bool.Parse(node.InnerText);
                            else if (node.Name == "run_at_startup")
                                CHBrun_at_startup.IsChecked = bool.Parse(node.InnerText);
                            else if (node.Name == "start_minimized")
                                CHBstart_minimized.IsChecked = bool.Parse(node.InnerText);
                            else if (node.Name == "minimize_to_tray")
                                CHBminimize_to_tray.IsChecked = bool.Parse(node.InnerText);
                            else if (node.Name == "auto_updates")
                                CHBauto_updates.IsChecked = auto_updates = bool.Parse(node.InnerText);
                            //added by version 1.3
                            else if (node.Name == "ss_voices")
                            {
                                if (node.InnerText != "null" && CBss_voices.Items.Count > 0)
                                    CBss_voices.SelectedItem = node.InnerText;
                            }
                            //added by version 1.7
                            else if (node.Name == "ss_volume")
                                TBss_volume.Text = node.InnerText;
                        }

                        //Checkboxes Checked and Unchecked events work only after form is loaded
                        //so they have to be called manually in order to load save data properly
                        CHBread_recognized_speech_Checked(new object(), new RoutedEventArgs());
                        CHBuse_better_dictation_Checked(new object(), new RoutedEventArgs());
                        CHBstart_with_hidden_Checked(new object(), new RoutedEventArgs());
                        CHBstart_minimized_Checked(new object(), new RoutedEventArgs());
                        CHBminimize_to_tray_Checked(new object(), new RoutedEventArgs());
                        CHBauto_updates_Checked(new object(), new RoutedEventArgs());
                        CHBsmart_mousegrid_Checked(new object(), new RoutedEventArgs());

                        set_values();
                    }
                }
                else if (File.Exists(old_file_path))
                {
                    load_settings_v_1_5_and_older();
                }
            }
            catch (Exception ex)
            {
                loading_error = true;
                MessageBox.Show(ex.Message, "Error AM016", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void fix_wrong_loaded_values()
        {
            try
            {
                if (confidence_turning_on > 99 || confidence_turning_on < 0)
                {
                    confidence_turning_on = 80;
                    TBconfidence_start.Text = confidence_turning_on.ToString();
                }
                if (confidence_other_commands > 99 || confidence_other_commands < 0)
                {
                    confidence_other_commands = 60;
                    TBconfidence_commands.Text = confidence_other_commands.ToString();
                }
                if (confidence_dictation > 99 || confidence_dictation < 0)
                {
                    confidence_dictation = 20;
                    TBconfidence_dictation.Text = confidence_dictation.ToString();
                }
                if (ss_volume > 100 || ss_volume < 0)
                {
                    ss_volume = 100;
                    TBss_volume.Text = ss_volume.ToString();
                }
                if (desired_figures_nr > max_figures_nr)
                {
                    desired_figures_nr = max_figures_nr;
                    TBdesired_figures_nr.Text = desired_figures_nr.ToString();
                }
                if (int.Parse(color_bg_str) < -16777216
                            || int.Parse(color_bg_str) > -1)
                {
                    color_bg_str = "-1973791";
                    color_bg = Color.FromRgb(225, 225, 225); //bg color
                }

                if (int.Parse(color_font_str) < -16777216
                    || int.Parse(color_font_str) > -1)
                {
                    color_font_str = "-16777216";
                    color_font = Color.FromRgb(0, 0, 0); //font color
                }
                if (font_size < 1 || font_size > max_font_size)
                {
                    font_size = 12;
                    TBfont_size.Text = font_size.ToString();
                    update_mousegrid_preview();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM017", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void load_coords()
        {
            FileStream fs = null;
            StreamReader sr = null;
            string file_path = System.IO.Path.Combine(Middle_Man.app_folder_path, filename_coords);

            try
            {
                if (File.Exists(file_path))
                {
                    fs = new FileStream(file_path, FileMode.Open, FileAccess.Read);
                    sr = new StreamReader(fs);

                    int x = 0, y = 0;
                    bool success1 = int.TryParse(sr.ReadLine(), out x);
                    bool success2 = int.TryParse(sr.ReadLine(), out y);

                    if (success1 && success2)
                    {
                        PositionSpeechWindow(x, y);
                    }

                    sr.Close();
                    fs.Close();
                }
                else
                    PositionSpeechWindow();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM018", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    if (sr != null)
                        sr.Close();
                    if (fs != null)
                        fs.Close();
                    PositionSpeechWindow();
                }
                catch (Exception ex2) { }
            }
        }

        void load_profiles()
        {
            Middle_Man.profiles = new List<Profile>();

            if (Directory.Exists(Middle_Man.profiles_path))
            {
                DirectoryInfo di = new DirectoryInfo(Middle_Man.profiles_path);

                FileInfo[] files = di.GetFiles("*.xml", SearchOption.TopDirectoryOnly);

                foreach (FileInfo file in files)
                {
                    load_profile(file.FullName);
                }

                Middle_Man.sort_profiles_by_name_asc();

                LVprofiles.Refresh();
            }
        }

        void load_profile(string path)
        {
            XmlDocument xml_doc = new XmlDocument();
            xml_doc.Load(path);

            XmlNodeList profile_tag = xml_doc.SelectNodes("//profile");

            int version = -1;
            bool parsing_v = false;

            if (profile_tag[0].Attributes["version"] != null)
                parsing_v = int.TryParse(profile_tag[0].Attributes["version"].Value, out version);

            //Work by Speech v. 1.5 and earlier had no version attribute (so we treat profiles made by these
            //versions as as version 1 of xml saving method for profiles and custom commands)

            if (version == -1 || (parsing_v && version == 1))
            {
                XmlNodeList nodes = xml_doc.SelectNodes("//profile")[0].ChildNodes;

                Profile p = new Profile();

                foreach (XmlNode node in nodes)
                {
                    if (node.Name == "name")
                        p.name = node.InnerText;
                    else if (node.Name == "program")
                        p.program = node.InnerText;
                    else if (node.Name == "enabled")
                        p.enabled = bool.Parse(node.InnerText);
                }

                p.custom_commands = new List<CustomCommand>();

                nodes = xml_doc.SelectNodes("//profile/custom_commands")[0].ChildNodes;

                int new_commands_nr = 0;

                foreach (XmlNode node in nodes)
                {
                    XmlNodeList nodes2 = node.ChildNodes;

                    CustomCommand cc = new CustomCommand();

                    foreach (XmlNode node2 in nodes2)
                    {
                        if (node2.Name == "name")
                            cc.name = node2.InnerText;
                        else if (node2.Name == "description")
                            cc.description = node2.InnerText;
                        else if (node2.Name == "group")
                            cc.group = node2.InnerText;
                        else if (node2.Name == "max_executions")
                            cc.max_executions = short.Parse(node2.InnerText);
                        else if (node2.Name == "enabled")
                            cc.enabled = bool.Parse(node2.InnerText);
                        else if (node2.Name == "actions")
                        {
                            cc.actions = new List<CC_Action>();

                            XmlNodeList actions_nodes = node2.ChildNodes;

                            foreach (XmlNode action in actions_nodes)
                            {
                                cc.actions.Add(new CC_Action()
                                {
                                    action = action.InnerText
                                });
                            }
                        }
                    }

                    p.custom_commands.Add(cc);
                    new_commands_nr++;
                }

                bool found = false;

                foreach (Profile profile in Middle_Man.profiles)
                {
                    if (profile.name == p.name)
                    {
                        found = true;
                        break;
                    }
                }

                if(found == false)
                    Middle_Man.profiles.Add(p);
            }
        }

        void compare_current_and_saved()
        {
            try
            {
                if (saving_enabled && dont_compare == false)
                {
                    SolidColorBrush scb = (SolidColorBrush)TBbackground_color.Background;
                    System.Drawing.Color color =
                        System.Drawing.Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
                    string current_TBbackground_color = color.ToArgb().ToString();

                    scb = (SolidColorBrush)TBfont_color.Background;
                    color =
                        System.Drawing.Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
                    string current_TBfont_color = color.ToArgb().ToString();

                    if (TBconfidence_start.Text != confidence_turning_on.ToString()
                    || TBconfidence_commands.Text != confidence_other_commands.ToString()
                    || TBconfidence_dictation.Text != confidence_dictation.ToString()
                    || CHBuse_better_dictation.IsChecked != better_dictation
                    || (CBss_voices.Items.Count > 0 &&
                       CBss_voices.SelectedItem.ToString() != ss_voice)
                    || TBss_volume.Text != ss_volume.ToString()
                    || CHBread_recognized_speech.IsChecked != read_recognized_speech
                    || CHBstart_with_hidden.IsChecked != start_with_hidden
                    || CHBrun_at_startup.IsChecked != run_at_startup
                    || CHBstart_minimized.IsChecked != start_minimized
                    || CHBminimize_to_tray.IsChecked != minimize_to_tray
                    || CHBauto_updates.IsChecked != auto_updates
                    || CBtype.SelectedItem.ToString() !=
                       grid_type.ToString().Replace("_", " ").FirstCharToUpper()
                    || CBlines.SelectedIndex != grid_lines
                    || current_TBbackground_color != color_bg_str
                    || current_TBfont_color != color_font_str
                    || TBdesired_figures_nr.Text != desired_figures_nr.ToString()
                    || TBfont_size.Text != font_size.ToString()
                    || CHBsmart_mousegrid.IsChecked != smart_grid)
                    {
                        if (Bsave_settings.IsEnabled == false)
                            Bsave_settings.IsEnabled = true;
                    }
                    else if (Bsave_settings.IsEnabled == true)
                        Bsave_settings.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM021", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void compare_current_and_default()
        {
            try
            {
                if (saving_enabled && dont_compare == false)
                {
                    SolidColorBrush scb = (SolidColorBrush)TBbackground_color.Background;
                    System.Drawing.Color color =
                        System.Drawing.Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
                    string current_TBbackground_color = color.ToArgb().ToString();

                    scb = (SolidColorBrush)TBfont_color.Background;
                    color =
                        System.Drawing.Color.FromArgb(scb.Color.A, scb.Color.R, scb.Color.G, scb.Color.B);
                    string current_TBfont_color = color.ToArgb().ToString();

                    if (TBconfidence_start.Text != default_confidence_turning_on.ToString()
                        || TBconfidence_commands.Text != default_confidence_other_commands.ToString()
                        || TBconfidence_dictation.Text != default_confidence_dictation.ToString()
                        || CHBuse_better_dictation.IsChecked != default_better_dictation
                        || (CBss_voices.Items.Count > 0 &&
                           CBss_voices.SelectedItem.ToString() != default_ss_voice)
                        || TBss_volume.Text != default_ss_volume.ToString()
                        || CHBread_recognized_speech.IsChecked != default_read_recognized_speech
                        || CHBstart_with_hidden.IsChecked != default_start_with_hidden
                        || CHBrun_at_startup.IsChecked != default_run_at_startup
                        || CHBstart_minimized.IsChecked != default_start_minimized
                        || CHBminimize_to_tray.IsChecked != default_minimize_to_tray
                        || CHBauto_updates.IsChecked != default_auto_updates
                        || CBtype.SelectedItem.ToString() !=
                           default_grid_type.ToString().Replace("_", " ").FirstCharToUpper()
                        || CBlines.SelectedIndex != default_grid_lines
                        || current_TBbackground_color != default_color_bg_str
                        || current_TBfont_color != default_color_font_str
                        || TBdesired_figures_nr.Text != default_desired_figures_nr.ToString()
                        || TBfont_size.Text != default_font_size.ToString()
                        || CHBsmart_mousegrid.IsChecked != default_smart_grid)
                    {
                        if (Brestore_default_settings.IsEnabled == false)
                            Brestore_default_settings.IsEnabled = true;
                    }
                    else if (Brestore_default_settings.IsEnabled == true)
                        Brestore_default_settings.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error AM022", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void register_first_run()
        {
            Microsoft.Win32.RegistryKey reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                    .OpenSubKey(Middle_Man.registry_path_easy, true);

            if (reg_key_easy == null)
            {
                reg_key_easy = Microsoft.Win32.Registry.CurrentUser.
                    CreateSubKey(Middle_Man.registry_path_easy, true);
                first_run = true;
            }
            reg_key_easy.SetValue(Middle_Man.registry_key_first_run, "yes");
        }

        string get_grid_folder_name_by_type(GridType gt)
        {
            if (gt == GridType.hexagonal)
                return folder_name_grid_hexagonal;
            else if (gt == GridType.square)
                return folder_name_grid_square;
            else if (gt == GridType.square_combined_precision)
                return folder_name_grid_combined;
            else if (gt == GridType.square_horizontal_precision)
                return folder_name_grid_horizontal;
            else if (gt == GridType.square_vertical_precision)
                return folder_name_grid_vertical;
            else return null;
        }

        private class MyWebClient : WebClient
        {
            protected override WebRequest GetWebRequest(Uri uri)
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                WebRequest w = base.GetWebRequest(uri);
                w.Timeout = 3000;
                return w;
            }
        }

        System.Diagnostics.Stopwatch watch;

        void start_time()
        {
            watch = System.Diagnostics.Stopwatch.StartNew();
        }

        long elapsed_ms = 0;

        void stop_time()
        {
            watch.Stop();
            elapsed_ms = watch.ElapsedMilliseconds;

            if (SW.TBrecognized_speech.IsInitialized)
            {
                SW.TBrecognized_speech.Dispatcher.Invoke(DispatcherPriority.Normal,
                    new Action(() => { SW.TBrecognized_speech.Text += ", " + elapsed_ms + " ms"; }));
            }
        }

        DateTime start_datetime;

        void start_time2()
        {
            start_datetime = DateTime.Now;
        }

        void stop_time2()
        {
            TimeSpan ts = DateTime.Now - start_datetime;
                        
            SW.TBrecognized_speech.Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() => { SW.TBrecognized_speech.Text += ", " + ts.TotalMilliseconds + " ms"; }));
        }

        Size MeasureString(string candidate, TextBlock textBlock)
        {
            var formattedText = new FormattedText(
                candidate,
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface(textBlock.FontFamily, textBlock.FontStyle, textBlock.FontWeight, textBlock.FontStretch),
                textBlock.FontSize,
                Brushes.Black,
                new NumberSubstitution(),
                1);

            return new Size(formattedText.Width, formattedText.Height);
        }
        
        public T DeepCopy<T>(T item)
        {
            BinaryFormatter formatter = new BinaryFormatter();
            MemoryStream stream = new MemoryStream();
            formatter.Serialize(stream, item);
            stream.Seek(0, SeekOrigin.Begin);
            T result = (T)formatter.Deserialize(stream);
            stream.Close();
            return result;
        }
    }

    public static class StringExtensions
    {
        public static string FirstCharToUpper(this string input)
        {
            switch (input)
            {
                case null: throw new ArgumentNullException(nameof(input));
                case "": throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input));
                default: return input[0].ToString().ToUpper() + input.Substring(1);
            }
        }
    }
}