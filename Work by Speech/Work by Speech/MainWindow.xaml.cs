//highest error nr: MW013
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Speech.Synthesis;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Vosk;
using WindowsInput;
using WindowsInput.Native;

//WaveIn_DataAvailable handles speech recognition
//ThreadPriority.Highest isn't really needed anywhere
//Fix cut window: x = -15px, y = -23px

namespace Speech
{
    public partial class MainWindow : Window
    {
        const string prog_version = "2.5";
              string latest_version = "";
        const string copyright_text = "Copyright © 2023 - 2026 Mikołaj Magowski. All rights reserved.";
        const string filename_model = "vosk-model-en-us-daanzu-20200905"; //Vosk speech recogniton model (7.08 (librispeech test-clean) 8.25 (tedlium))
        const string filename_settings = "settings.xml";
        const string filename_coords = "coords.txt"; //speech recognition window last location
        const string grids_foldername = "grids";
        const int grid_symbols_limit = 50;//50-58 recommended
        const int max_font_size = 400;
        const bool resized_grid = true; //when true, resizes mousegrid so screen is fully covered
        const bool movable_grid = true; //when true, mousegrid can be moved by speech

        //Mode names in change mode button (in SpeechWindow) and notify icon (ni) context menu:
        const string mode_off = "OFF";
        const string mode_command = "Command";
        const string mode_dictation = "Dictation";

        const string icon_off = "pack://application:,,,/Work by Speech;component/images/off2.ico";
        const string icon_command = "pack://application:,,,/Work by Speech;component/images/command2.ico";
        const string icon_dictation = "pack://application:,,,/Work by Speech;component/images/dictation2.ico";

        const string folder_name_grid_hexagonal = "hexagonal";
        const string folder_name_grid_square = "square";
        const string folder_name_grid_horizontal = "square_horizontal";
        const string folder_name_grid_vertical = "square_vertical";
        const string folder_name_grid_combined = "square_combined";

        mode current_mode = mode.off;
        mode last_mode = mode.command; //before turning off

        bool saving_enabled = false;
        bool loading_error = false;

        Process prc;

        List<string> list_current; //most commands available in current mode
                                   //(not apps switching and opening commands, not any app and foreground app custom commands, not mousegrid commands)
        List<string> list_switch_to_apps;
        List<string> list_open_apps;
        List<string> list_cc_foreground; //custom commands foreground program
        List<string> list_cc_any; //custom commands any program
        List<string> list_mousegrid;
        //------------------------------
        List<string> list_off_mode;
        List<string> list_dictation; 
        List<string> list_builtin_commands;        

        List<string> dictation_commands;

        //---------------Main voice commands START---------------
        const string turn_on = "start speech recognition";
        const string turn_off = "recognition off";
        const string switch_to_command_mode = "command";
        const string switch_to_dictation_mode = "dictation";
        const string show_speech_recognition = "show speech recognition"; //window
        const string hide_speech_recognition = "hide speech recognition";
        const string open_app_str = "Open"; //capital O here!!!
        const string switch_to_app_str = "switch to";

        string[] cancels_str = new string[] { "cancel", "escape" };
        string[] directions_str = new string[] { "up", "down", "left", "right" }; //for moving mousegrid

        string[] s_combo = new string[] { "control", "shift", "alt", "windows" };
        string[] s_combo2 = new string[] { "control", "shift", "alt", "windows" };
        string[] s_mouse_moves = new string[] { "up", "down", "left", "right" };
        //string[] s_mouse_moves = new string[] { "move up", "move down", "move left", "move right" };
        string[] drag_edges_str = new string[] { "top edge", "bottom edge", "left edge", "right edge",
            "screen center"};
        string[] move_edges_str = new string[] { "top edge", "bottom edge", "left edge", "right edge",
            "screen center"};
        string[] s_other = new string[] { switch_to_dictation_mode, show_speech_recognition,
            hide_speech_recognition};
        //---------------Main voice commands END-----------------

        InputSimulator sim = new InputSimulator();

        Thread THRswitch_to, THRmouse, THRcommands, THRmonitor;
        Thread THRholder; //for holding keys
        //bool thread_abort1 = false; //redeclaring and starting stopped thread doesn't work (weird)
        bool recognition_suspended = false;
        bool first_run = false;

        List<Grid_Symbol> grid_alphabet = new List<Grid_Symbol>();
        List<VirtualKeyCode> keys_to_hold = new List<VirtualKeyCode>();

        //Used by speech synthesis:
        List<string> ss_voices_priority_list = new List<string>() { "Zira", "Susan", "Hazel", "Linda", 
            "Catherine", "Heera", "Sean" };
        //Zira - US, Susan/Hazel - UK, Linda - Canada, Catherine - Australia, Heera - India, Sean - Ireland
        IReadOnlyCollection<InstalledVoice> installed_voices;

        //---DEFAULT SETTINGS START---:
        const int default_confidence_turning_on = 80;
        const int default_confidence_other_commands = 60;
        const int default_confidence_dictation = 20;
        
        string default_ss_voice = "";
        const int default_ss_volume = 100;
        const bool default_read_recognized_speech = false;

        const bool default_start_with_hidden = false;
        const bool default_run_at_startup = false;
        const bool default_start_minimized = false;
        const bool default_minimize_to_tray = false;
        const bool default_auto_updates = true;

        const GridType default_grid_type = GridType.square_horizontal_precision;
        const int default_grid_lines = 0; //0-2 (0 - no lines, 1 - dotted lines, 2 - normal lines)
        const int default_desired_figures_nr = 2550;
        const string default_color_bg_str = "-1973791"; //light grey
        const string default_color_font_str = "-16777216"; //black
        const int default_font_size = 12;
        const bool default_smart_grid = true;
        //---DEFAULT SETTINGS END---

        //---USER SETTINGS START---
        int confidence_turning_on;
        int confidence_other_commands;

        string ss_voice;
        int ss_volume;
        bool read_recognized_speech;

        bool start_with_hidden;
        bool run_at_startup;
        bool start_minimized;
        bool minimize_to_tray;
        bool auto_updates;

        GridType grid_type;
        int grid_lines; //0-2 (0 - no lines, 1 - dotted lines, 2 - normal lines)
        int desired_figures_nr;
        string color_bg_str; //light grey
        string color_font_str; //black        
        int font_size;
        bool smart_grid;
        //---USER SETTINGS END---

        bool apps_switching = true;
        bool apps_opening = true;

        int max_figures_nr = 1; //for full alphabet
        int max_figures; //for decreased alphabet (decreased so whole mousegrid can be covered)

        //------------Mousegrid moving by speech settings START------------
        double offset = 0.25; //offset by percentage of figure height
        int offset_x = 0; //grid offset
        int offset_y = 0;
        //------------Mousegrid moving by speech settings END--------------

        Color color_font;// = Color.FromRgb(0, 0, 0); //font color
        //Color color_bg = Color.FromRgb(255, 255, 255);
        Color color_bg;// = Color.FromRgb(225, 225, 225); //bg color

        //bool auto_grid_font_size = true; //bad idea
        FontFamily font_family = new FontFamily("Verdana");
        //FontFamily font_family = new FontFamily("Tahoma");
        //FontFamily font_family = new FontFamily("Microsoft Sans Serif");
        //FontFamily font_family = new FontFamily("Calibri");//12,10
        //FontFamily font_family = new FontFamily("Arial");
        //FontFamily font_family = new FontFamily("Times New Roman");
        //FontFamily font_family = new FontFamily("Courier New");
        
        int screen_width, screen_height;

        List<int> rows = new List<int>();
        List<int> cols = new List<int>();

        int rows_nr, cols_nr;

        int figures;
        double figure_width, figure_height;
        double grid_width, grid_height;

        MouseGrid MW;
        bool grid_visible = false;
        bic_type last_command = bic_type.cancel;

        List<Process_grid> grids = new List<Process_grid>();
        List<Installed_App> installed_apps = new List<Installed_App>();
        int grid_ind = 0;
        string start_menu_path = "";
        
        SpeechWindow SW;

        bool inside_speech_recognized_event = false;

        WaveInEvent waveIn;
        static VoskRecognizer recognizer;

        public SpeechSynthesizer ss = new SpeechSynthesizer();

        System.Windows.Forms.MenuItem mi_switch_to_command_mode;
        System.Windows.Forms.MenuItem mi_switch_to_dictation_mode;
        System.Windows.Forms.MenuItem mi_switch_to_off_mode;
        System.Windows.Forms.MenuItem mi_exit;

        public MainWindow()
        {
            if (test_mode > 0)
            {
                if (test1_on > 0)
                    test1();

                if (test2_on > 0)
                    test2();

                if (test3_on > 0)
                    test3();

                if (test4_on > 0)
                    test4();

                if (test5_on > 0)
                    test5();
            }

            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                saving_enabled = false;
                Middle_Man.saving_folder_path = Path.Combine(Middle_Man.users_directory_path,
                    Middle_Man.prog_name);
                Middle_Man.profiles_path = Path.Combine(Middle_Man.saving_folder_path,
                    Middle_Man.profiles_foldername);

                prc = Process.GetCurrentProcess();
                prc.PriorityClass = ProcessPriorityClass.High;

                Stream iconStream = System.Windows.Application.GetResourceStream(
                    new Uri(icon_off)).Stream;
                ni.Icon = new System.Drawing.Icon(iconStream);
                iconStream.Close();
                ni.MouseClick += new System.Windows.Forms.MouseEventHandler(ni_MouseClick);
                ni.Visible = true;

                System.Windows.Forms.ContextMenu cm = new System.Windows.Forms.ContextMenu();
                mi_switch_to_command_mode = new System.Windows.Forms.MenuItem();
                mi_switch_to_dictation_mode = new System.Windows.Forms.MenuItem();
                mi_switch_to_off_mode = new System.Windows.Forms.MenuItem();
                mi_exit = new System.Windows.Forms.MenuItem();

                // Initialize contextMenu1
                cm.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { mi_switch_to_command_mode });
                cm.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { mi_switch_to_dictation_mode });
                cm.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { mi_switch_to_off_mode });
                cm.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] { mi_exit });

                mi_switch_to_command_mode.Index = 0;
                mi_switch_to_command_mode.Text = "Switch to " + mode_command + " mode";
                mi_switch_to_command_mode.Click += new System.EventHandler(mi_switch_to_command_mode_Click);

                mi_switch_to_dictation_mode.Index = 1;
                mi_switch_to_dictation_mode.Text = "Switch to " + mode_dictation + " mode";
                mi_switch_to_dictation_mode.Click += new System.EventHandler(mi_switch_to_dictation_mode_Click);

                mi_switch_to_off_mode.Index = 2;
                mi_switch_to_off_mode.Text = "Switch to " + mode_off + " mode";
                mi_switch_to_off_mode.Click += new System.EventHandler(mi_switch_to_off_mode_Click);

                mi_exit.Index = 3;
                mi_exit.Text = "Exit";
                mi_exit.Click += new System.EventHandler(mi_exit_Click);

                ni.ContextMenu = cm;

                StringBuilder path = new StringBuilder(260);

                try
                {
                    //Access All Users Start Menu
                    SHGetSpecialFolderPath(IntPtr.Zero, path, CSIDL_COMMON_STARTMENU, false);
                    start_menu_path = path.ToString();

                    Directory.GetFiles(start_menu_path, "*.*", SearchOption.AllDirectories);
                }
                catch (Exception not_used)
                {
                    //Directory.GetFiles may throw access denied expection if some folders in
                    //"C:\\ProgramData\\Microsoft\\Windows\\Start Menu" path that it returns have denied access

                    //In non-US Windows 11 installations Directory.GetFiles for "C:\\ProgramData\\Microsoft\\Windows\\Start Menu" path
                    //returns path to 'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs' directory with 'Programs'
                    //named in Windows installation language, but this directory doesn't exist
                    //(it's a Windows 11 bug, [possibly Windows 10 too] which occurs after changing Windows language)

                    //This solves 2 above problems:
                    start_menu_path = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs";

                    try
                    {
                        Directory.GetFiles(start_menu_path, "*.*", SearchOption.AllDirectories);
                    }
                    catch (Exception not_used2)
                    {
                        try
                        {
                            start_menu_path = path.ToString();
                            Directory.GetFiles(start_menu_path, "*.*", SearchOption.AllDirectories);
                        }
                        catch (Exception ex)
                        {
                            start_menu_path = "";

                            Microsoft.Win32.RegistryKey reg_key_easy = Microsoft.Win32.Registry.CurrentUser
                                    .OpenSubKey(Middle_Man.registry_path_easy, true);

                            if (reg_key_easy == null)
                            {
                                MessageBox.Show(ex.Message + "\n\n" +
                                    "This error means that you can't use 'Open app_name' speech command. ",
                                    "Error MW004a", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                            else
                            {
                                string display_error_MW004a = "";

                                object ob = reg_key_easy.GetValue(Middle_Man.registry_key_display_error_MW004a);

                                if (ob != null)
                                {
                                    display_error_MW004a = ob.ToString();
                                }

                                if (display_error_MW004a != "No")
                                {
                                    MessageBoxResult mbr = MessageBox.Show(ex.Message + "\n\n" +
                                        "This error means that you can't use 'Open app_name' speech command. " +
                                        "Click 'Yes' if you want this error to appear the next time you run Work by Speech.",
                                        "Error MW004a", MessageBoxButton.YesNo, MessageBoxImage.Error);

                                    if (mbr == MessageBoxResult.No)
                                    {
                                        reg_key_easy.SetValue(Middle_Man.registry_key_display_error_MW004a, "No");
                                    }
                                }
                            }
                        }
                    }
                }

                //not enough shortcuts:
                //start_menu_path = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);

                is_program_already_running();

                InitializeComponent();

                this.Title = Middle_Man.prog_name + " " + prog_version;
                Lprogram_name.Content = Middle_Man.prog_name;
                Linstalled_version.Content = "Installed version: " + prog_version;
                Lhomepage.Content = Middle_Man.url_homepage;
                Lcopyright.Content = copyright_text;

                Bedit_profile.IsEnabled = false;
                Bdelete_profile.IsEnabled = false;
                Badd_command.IsEnabled = false;
                Bedit_command.IsEnabled = false;
                Bedit_action.IsEnabled = false;
                Bdelete_command.IsEnabled = false;
                Bexport_profiles.IsEnabled = false;

                MIenable_profiles.IsEnabled = false;
                MIdisable_profiles.IsEnabled = false;
                MIedit_profiles.IsEnabled = false;
                MIdelete_profiles.IsEnabled = false;
                MIduplicate_profiles.IsEnabled = false;
                MIenable_commands.IsEnabled = false;
                MIdisable_commands.IsEnabled = false;
                MIedit_commands.IsEnabled = false;
                MIdelete_commands.IsEnabled = false;
                MIcopy_commands.IsEnabled = false;
                MIpaste_commands.IsEnabled = false;
                MIedit_actions.IsEnabled = false;

                Benable_bic_off.IsEnabled = false;
                Bdisable_bic_off.IsEnabled = false;
                Benable_bic_general.IsEnabled = false;
                Bdisable_bic_general.IsEnabled = false;
                Benable_bic_mouse.IsEnabled = false;
                Bdisable_bic_mouse.IsEnabled = false;
                Benable_bic_pressing.IsEnabled = false;
                Bdisable_bic_pressing.IsEnabled = false;
                Benable_bic_inserting.IsEnabled = false;
                Bdisable_bic_inserting.IsEnabled = false;
                Benable_bic_dictation.IsEnabled = false;
                Bdisable_bic_dictation.IsEnabled = false;

                MIenable_bic_off.IsEnabled = false;
                MIdisable_bic_off.IsEnabled = false;
                MIenable_bic_general.IsEnabled = false;
                MIdisable_bic_general.IsEnabled = false;
                MIenable_bic_mouse.IsEnabled = false;
                MIdisable_bic_mouse.IsEnabled = false;
                MIenable_bic_pressing.IsEnabled = false;
                MIdisable_bic_pressing.IsEnabled = false;
                MIenable_bic_inserting.IsEnabled = false;
                MIdisable_bic_inserting.IsEnabled = false;
                MIenable_bic_dictation.IsEnabled = false;
                MIdisable_bic_dictation.IsEnabled = false;

                installed_voices = ss.GetInstalledVoices();

                foreach (InstalledVoice iv in installed_voices)
                {
                    CBss_voices.Items.Add(iv.VoiceInfo.Name);
                }

                foreach (GridType type in (GridType[])Enum.GetValues(typeof(GridType)))
                {
                    CBtype.Items.Add(type.ToString().Replace("_", " ").FirstCharToUpper());
                }

                CBlines.Items.Add("None");
                CBlines.Items.Add("Dotted");
                CBlines.Items.Add("Solid");

                TBbackground_color.IsReadOnly = true;
                TBfont_color.IsReadOnly = true;

                TBmousegrid_preview.Text = "qg ƒf -: °\" li1";
                TBmousegrid_preview.FontFamily = font_family;
                TBmousegrid_preview.TextAlignment = TextAlignment.Center;
                TBmousegrid_preview.VerticalAlignment = VerticalAlignment.Center;

                max_figures_nr = (int)Math.Pow((double)grid_symbols_limit, 2) + grid_symbols_limit;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW001", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            register_first_run();

            try
            {
                try
                {
                    string content;
                    MyWebClient wc = new MyWebClient();
                    content = wc.DownloadString(Middle_Man.url_latest_version);

                    latest_version = content.Replace("\r\n", "").Trim();
                }
                catch (WebException we)
                {
                    latest_version = "unknown";
                }

                Llatest_version.Content = "Latest version: " + latest_version;

                restore_default_settings();

                set_values(); //important

                load_settings();

                fix_wrong_loaded_values();

                saving_enabled = true;

                update_mousegrid_preview();

                compare_current_and_default();
                compare_current_and_saved();

                if (loading_error)
                {
                    MessageBox.Show("Loading error was detected. All settings" +
                        " will be restored to default and saved.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    restore_default_settings();
                    save_settings(); //save settings so loading error won't happen again (default values
                                     //will take place of unread values)
                }

                CenterWindowOnScreen();

                SW = new SpeechWindow();
                SW.Topmost = true;
                SW.ShowInTaskbar = false;

                load_coords();

                if (test_mode > 0)
                {
                    SW.TBrecognized_speech.Text = "TEST MODE";
                    SW.TBrecognized_speech.FontSize = 24;

                    this.Title += " - TEST MODE";
                }

                if (CHBstart_minimized.IsChecked == true)
                {
                    WindowState = WindowState.Minimized;

                    if (CHBminimize_to_tray.IsChecked == true)
                    {
                        this.Hide();
                    }
                }

                // Configure speech recognizer.
                Vosk.Vosk.SetLogLevel(-1); // You can set to -1 to disable logging messages

                string modelPath = @"vosk-model-en-us-daanzu-20200905"; //ma potencjał (zdobył 197 punktów)

                Model model = new Model(modelPath);

                recognizer = new VoskRecognizer(model, 16000.0f);

                waveIn = new WaveInEvent
                {
                    DeviceNumber = 0,
                    WaveFormat = new WaveFormat(16000, 1),
                    BufferMilliseconds = 250
                };

                waveIn.DataAvailable += WaveIn_DataAvailable;
                
                try
                {
                    waveIn.StartRecording();
                }
                catch (Exception ex5)
                {
                    MessageBox.Show("Please connect a microphone to your computer before running this " +
                        "application.",
                            "Error MW003", MessageBoxButton.OK, MessageBoxImage.Error);
                    Process.GetCurrentProcess().Kill();
                }

                load_profiles();

                Middle_Man.load_groups();

                LVprofiles.ItemsSource = Middle_Man.profiles;
                cv_LVprofiles = (CollectionView)CollectionViewSource.GetDefaultView(
                        LVprofiles.ItemsSource);

                create_bic_lists();

                load_bic_toggling_data();

                if (is_bic_in_general_and_mouse_enabled(bic_type.switch_to_app))
                    apps_switching = true;
                else
                    apps_switching = false;

                if (is_bic_in_general_and_mouse_enabled(bic_type.open_app))
                    apps_opening = true;
                else
                    apps_opening = false;

                create_lists();

                load_turned_off();

                if (smart_grid)
                    load_grids();

                if (CHBstart_with_hidden.IsChecked == false)
                {
                    SW.Show();
                }

                Middle_Man.force_updating_both_cc_lists = true;

                single_app_lists_update();

                get_installed_apps();

                THRmonitor = new Thread(new ThreadStart(monitor));
                THRmonitor.Start();

                THRholder = new Thread(new ThreadStart(hold_keys_and_buttons));
                THRholder.Start();

                THRswitch_to = new Thread(new ThreadStart(keep_lists_updated));
                THRswitch_to.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW004", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            Mouse.OverrideCursor = null;
        }
        
        //List<string> list_custom_commands_foreground; //foreground program
        //List<string> list_apps_switching;
        //List<string> list_apps_opening;

        void create_lists()
        {
            list_off_mode = new List<string>();
            list_builtin_commands = new List<string>();
            list_dictation = new List<string>();

            if (are_all_bic_off_disabled() == false)
            {
                list_off_mode = create_off_mode_list();
            }

            if (are_all_bic_general_and_mouse_disabled() == false)
            {
                list_builtin_commands = create_builtin_commands_list();
            }

            if (are_all_bic_dictation_disabled() == false)
            {
                list_dictation = create_dictation_commands_list();
            }
            
            list_mousegrid = create_grid_list(true);

            list_cc_any = create_custom_commands_list(Middle_Man.any_program_name);
        }

        void is_program_already_running()
        {
            Process[] arr = Process.GetProcesses();
            string[] a;
            int i = 0;

            foreach (Process p in arr)
            {
                if(p.ProcessName == Middle_Man.prog_name)
                {
                    i++;
                }    
            }

            if(i > 1)
            {
                MessageBox.Show(Middle_Man.prog_name + " is already running.",
                        "Error MW007", MessageBoxButton.OK, MessageBoxImage.Error);
                Process.GetCurrentProcess().Kill();
            }
        }

        bool get_installed_apps()
        {
            if (string.IsNullOrEmpty(start_menu_path))
                return false;

            string[] allfiles = Directory.GetFiles(start_menu_path, "*.*", SearchOption.AllDirectories);
            
            string[] a;
            string str;
            List<string> names;
            string installed_app_path, s;
            int name_length;
            Installed_App installed_app;

            List<Installed_App> installed_apps2 = new List<Installed_App>();

            foreach (var file in allfiles)
            {
                FileInfo info = new FileInfo(file);
            }

            foreach (var file in allfiles)
            {
                FileInfo info = new FileInfo(file);
                if (info.Name.ToLower().Contains("install")
                    || info.Name.ToLower().Contains("application") || info.Name.Length < 6
                    || info.Name.ToLower().Substring(info.Name.Length - 4) != ".lnk")
                    continue;

                installed_app = new Installed_App();
                installed_app.names = new List<string>();

                a = info.Name.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                s = a[0];

                for (int i = 1; i < a.Length - 1; i++)
                {
                    s += "." + a[i];
                }

                s = s.Trim();
                a = s.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < a.Length; i++)
                {
                    a[i] = a[i].Trim();
                    if (a[i].Length > 1)
                        installed_app.names.Add(a[i]);
                }
                for (int i = 0; i < a.Length - 1; i++)
                {
                    str = a[i].Trim() + " " + a[i + 1].Trim();
                    if (str.Length > 3)
                    {
                        installed_app.names.Add(a[i].Trim() + " " + a[i + 1].Trim());
                    }
                }
                if (a.Length > 2)
                {
                    if(s.Length > 3)
                        installed_app.names.Add(s);
                }

                installed_app.name_length = info.Name.Length;
                installed_app.path = info.FullName;
                installed_apps2.Add(installed_app);
            }

            //sort apps by app name length ascending
            for (int i = 0; i < installed_apps2.Count; i++)
            {
                for (int j = 0; j < installed_apps2.Count; j++)
                {
                    if (installed_apps2[j].name_length > installed_apps2[i].name_length)
                    {
                        names = installed_apps2[j].names;
                        installed_app_path = installed_apps2[j].path;
                        name_length = installed_apps2[j].name_length;
                        installed_apps2[j].names = installed_apps2[i].names;
                        installed_apps2[j].path = installed_apps2[i].path;
                        installed_apps2[j].name_length = installed_apps2[i].name_length;
                        installed_apps2[i].names = names;
                        installed_apps2[i].path = installed_app_path;
                        installed_apps2[i].name_length = name_length;
                    }
                }
            }

            installed_apps = installed_apps2;

            //dt1 = DateTime.Now;
            //dt2 = DateTime.Now;
            //ts = dt2 - dt1;
            //MessageBox.Show("processes: " + ts.TotalMilliseconds.ToString());

            //if get_installed_apps() == true means apps number has changed
            list_open_apps = new List<string>();

            s = open_app_str;

            if (test_mode > 0 && test3_on > 0)
            {
                test3();
            }
            else
            {
                foreach (Installed_App app in installed_apps)
                {
                    foreach (string name2 in app.names)
                    {
                        if (name2 != null && name2 != "")
                        {
                            list_open_apps.Add(s + " " + name2);
                        }
                    }
                }
            }

            return true;
        }

        void load_turned_off(bool loaded_from_speech_recognized = false)
        {
            int i = 0;
            
            recognition_suspended = true;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Hidden;
            }));

            //THRswitch_to.Abort(); //abort only causes problems

            //wait for suspend:
            while (inside_speech_recognized_event == true && loaded_from_speech_recognized == false)
            {
                Thread.Sleep(10);
                i++;
            }

            list_current = new List<string>();
            add_to_list_current(list_type.list_off_mode);

            Stream iconStream = System.Windows.Application.GetResourceStream(
                new Uri(icon_off)).Stream;

            //it's safer to keep recognition_suspended = false; here (to avoid possible
            //freeze if a user would say "Recognition Off" while clicking save settings)
            recognition_suspended = false;

            SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.mode = 0;
                SW.Bmode.Content = mode_off;
                SW.Bmode.Foreground = new SolidColorBrush(Color.FromRgb(232, 4, 4));

                // Create a BitmapSource  
                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(icon_off);
                bitmap.EndInit();

                SW.Icon = bitmap;
            }));

            ni.Icon = new System.Drawing.Icon(iconStream);
            iconStream.Close();

            mi_switch_to_off_mode.Visible = false;
            mi_switch_to_dictation_mode.Visible = true;
            mi_switch_to_command_mode.Visible = true;

            current_mode = mode.off;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Visible;
            }));
        }

        void load_turned_on(bool loaded_from_speech_recognized = false)
        {
            int i = 0;
            
            recognition_suspended = true;

            /* Removing this solved freeze problem when using west command to save settings
            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Hidden;
            }));
            */

            //THRswitch_to.Abort(); //abort only causes problems
            
            //wait for suspend:
            while (inside_speech_recognized_event == true && loaded_from_speech_recognized == false)
            {
                Thread.Sleep(10);
                i++;
            }

            if (current_mode == mode.command)
            {
                list_current = new List<string>();
                add_to_list_current(list_type.list_builtin_commands);

                lock (lock_list_cc_foreground)
                {
                    list_cc_foreground = new List<string>();
                }

                Stream iconStream = System.Windows.Application.GetResourceStream(
                        new Uri(icon_command)).Stream;

                //we must set recognition_suspended to false, before using invoke,
                //otherwise using west command to save settings (or clicking save settings
                //button while a speech recognition action is performed) would cause a freeze
                //(save settings thread waits for recognition_suspended == false)
                recognition_suspended = false;

                SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                {
                    SW.mode = 1;
                    SW.Bmode.Content = mode_command;
                    SW.Bmode.Foreground = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                    
                    // Create a BitmapSource  
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(icon_command);
                    bitmap.EndInit();
                    
                    SW.Icon = bitmap;
                }));

                ni.Icon = new System.Drawing.Icon(iconStream);
                iconStream.Close();

                mi_switch_to_off_mode.Visible = true;
                mi_switch_to_dictation_mode.Visible = true;
                mi_switch_to_command_mode.Visible = false;

                current_mode = mode.command;
                last_mode = current_mode;
            }
            else if(current_mode == mode.dictation)
            {
                list_current = new List<string>();
                add_to_list_current(list_type.list_dictation);

                Stream iconStream = System.Windows.Application.GetResourceStream(
                new Uri(icon_dictation)).Stream;

                //we must set recognition_suspended to false, before using invoke,
                //otherwise using west command to save settings (or clicking save settings
                //button while a speech recognition action is performed) would cause a freeze
                //(save settings thread waits for recognition_suspended == false)
                recognition_suspended = false;

                SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                {
                    SW.mode = 2;
                    SW.Bmode.Content = mode_dictation;
                    SW.Bmode.Foreground = new SolidColorBrush(Color.FromRgb(0, 88, 255));

                    // Create a BitmapSource  
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(icon_dictation);
                    bitmap.EndInit();

                    SW.Icon = bitmap;
                }));

                ni.Icon = new System.Drawing.Icon(iconStream);
                iconStream.Close();

                mi_switch_to_off_mode.Visible = true;
                mi_switch_to_dictation_mode.Visible = false;
                mi_switch_to_command_mode.Visible = true;

                current_mode = mode.dictation;
                last_mode = current_mode;
            }

            recognition_suspended = false;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Visible;
            }));
        }

        private void enable_grid()
        {
            int i = 0;
            
            recognition_suspended = true;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Hidden;
            }));

            //THRswitch_to.Abort(); //abort only causes problems

            //wait for suspend:
            while (inside_speech_recognized_event == true)
            {
                Thread.Sleep(10);
                i++;
            }

            current_mode = mode.grid;

            recognition_suspended = false;
        }

        string s_last_ch = "";
        void keep_lists_updated()
        {
            while (true)
            {
                try
                {
                    if (current_mode == mode.command)
                    {
                        single_app_lists_update();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error MW008", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                Thread.Sleep(1000);
            }
        }

        string last_foreground_window_process = "";
        bool inside_app_lists_update = false;

        void single_app_lists_update()
        {
            inside_app_lists_update = true;

            string s = "", w;

            DateTime dt1 = DateTime.Now, dt2 = DateTime.Now;
            TimeSpan ts;

            Process[] arr = null;
            string[] a;
            string name;

            arr = Process.GetProcesses();
            
            IntPtr handle = GetForegroundWindow();
            string process_name = "";

            foreach (Process p in arr)
            {
                if (p.MainWindowHandle == handle)
                {
                    process_name = p.ProcessName;
                    break;
                }
            }

            process_name = process_name.ToLower();
            List<string> new_list_cc_foreground = create_custom_commands_list(process_name);

            if (Middle_Man.force_updating_both_cc_lists)
            {
                lock (lock_list_cc_foreground)
                {
                    list_cc_foreground = new_list_cc_foreground;
                }

                lock (lock_list_cc_any)
                {
                    list_cc_any = create_custom_commands_list(Middle_Man.any_program_name);
                }

                last_foreground_window_process = process_name;

                Middle_Man.force_updating_both_cc_lists = false;
            }
            else if (last_foreground_window_process != process_name)
            {
                lock (lock_list_cc_foreground)
                {
                    list_cc_foreground = new_list_cc_foreground;
                }

                last_foreground_window_process = process_name;
            }

            if (apps_switching)
            {
                lock (lock_list_switch_to_apps)
                {
                    list_switch_to_apps = new List<string>();

                    s = switch_to_app_str;

                    //around 150ms on my CPU (Q9550)
                    foreach (Process p in arr)
                    {
                        name = p.MainWindowTitle.Replace("\\\"", "").Replace("\"", "");
                        name = name.Trim();

                        //is process window open in taskbar?
                        if (name != null && name.Length > 0)
                        {
                            a = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                            for (int i = 0; i < a.Length; i++)
                            {
                                a[i] = a[i].Trim();

                                if (a[i].Length > 1)
                                {
                                    list_switch_to_apps.Add(s + " " + a[i]);
                                }
                            }
                            for (int i = 0; i < a.Length - 1; i++)
                            {
                                a[i] = a[i].Trim();
                                a[i + 1] = a[i + 1].Trim();

                                if (a[i].Length > 0 || a[i + 1].Length > 0)
                                {
                                    w = a[i] + " " + a[i + 1];
                                    list_switch_to_apps.Add(s + " " + w);
                                }
                            }
                            if (a.Length > 2)
                            {
                                list_switch_to_apps.Add(s + " " + name);
                            }
                        }
                    }
                }
            }

            inside_app_lists_update = false;
        }

        List<string> create_off_mode_list()
        {
            List<string> off_mode_list = new List<string>();

            if(are_all_bic_off_disabled() == false)
                off_mode_list.Add(turn_on);

            return off_mode_list;
        }

        List<string> create_dictation_commands_list()
        {
            List<string> dictation_commands_list = new List<string>();

            for (int i = 0; i < list_bic_dictation.Count(); i++)
            {
                if (list_bic_dictation[i].enabled)
                {
                    dictation_commands_list.Add(list_bic_dictation[i].name);
                }
            }

            return dictation_commands_list;
        }

        List<string> create_grid_list(bool reset_smart_grid)
        {
            List<string> grid_list = new List<string>();

            //create_full_grid_alphabet();
            create_optimized_grid_alphabet();
            //create_normal_grid_alphabet();
            //create_wide_grid_alphabet();

            //Mathematically incorrect, but returns desired value (thanks to Floor):
            int count = (int)Math.Floor(Math.Sqrt((double)desired_figures_nr));

            string w, s;

            //sort by word length
            for (int i = 0; i < count - 1; i++)
            {
                for (int j = i + 1; j < count; j++)
                {
                    if (grid_alphabet[j].word.Length < grid_alphabet[i].word.Length)
                    {
                        w = grid_alphabet[j].word;
                        s = grid_alphabet[j].symbol;
                        grid_alphabet[j].word = grid_alphabet[i].word;
                        grid_alphabet[j].symbol = grid_alphabet[i].symbol;
                        grid_alphabet[i].word = w;
                        grid_alphabet[i].symbol = s;
                    }
                }
            }

            if (count > grid_symbols_limit)
                count = grid_symbols_limit;

            max_figures = count + count * count;

            //returns scaled width and height if windows scaling is enabled in system settings
            screen_width = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenWidth);
            screen_height = (int)Math.Round(System.Windows.SystemParameters.PrimaryScreenHeight);

            //returns real width and height if windows scaling is enabled in system settings
            //screen_width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
            //screen_height = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;

            rows = new List<int>();
            cols = new List<int>();

            int a, b;

            //i - square/hexagon width
            for (int i = 1; i < 2000; i++)
            {
                a = (int)Math.Round((double)(screen_width / i));
                b = (int)Math.Round((double)(screen_height / i));

                if (grid_type == GridType.hexagonal)
                {
                    if ((a * b + (a - 1) * (b - 1)) <= max_figures)
                    {
                        cols.Add(a);
                        rows.Add(b);
                    }
                }
                else
                {
                    if (a * b <= max_figures)
                    {
                        cols.Add(a);
                        rows.Add(b);
                    }
                }
            }

            rows_nr = rows[0];
            cols_nr = cols[0];

            if (grid_type == GridType.hexagonal)
                figures = cols_nr * rows_nr + (cols_nr - 1) * (rows_nr - 1);
            else
                figures = cols_nr * rows_nr;

            figure_width = screen_width / (double)cols_nr;
            figure_height = screen_height / (double)rows_nr;

            if (resized_grid == true)
            {
                grid_width = screen_width;
                grid_height = screen_height;
            }
            else
            {
                grid_width = figure_width * (double)cols_nr;
                grid_height = figure_height * (double)rows_nr;
            }

            //bad idea:
            //if (auto_grid_font_size)
            //{
            //    if (grid_type == GridType.hexagonal)
            //        font_size = (int)(Math.Floor(figure_width * 14 / 35.555555555555557));
            //    else
            //        font_size = (int)(Math.Floor(figure_width * 14 / 25.098039215686274));
            //}

            if (reset_smart_grid)
            {
                grids = new List<Process_grid>();

                //"Default Process_grid ind=0" is used to copy default grid elements
                //to new grids (for different apps)
                grids.Add(new Process_grid("Default Process_grid ind=0"));

                int ind = 0;
                int count2;

                count2 = (int)Math.Ceiling(Math.Sqrt(figures));
                if (count2 < count)
                    count = count2;

                for (int j = 0; j < count && ind < figures; j++)
                {
                    grids[0].elements.Add(
                        new Grid_element(grid_alphabet[j].word.Replace(" ", ""), grid_alphabet[j].symbol));

                    ind++;
                }

                for (int i = 0; i < count && ind < figures; i++)
                {
                    for (int j = 0; j < count && ind < figures; j++)
                    {
                        grids[0].elements.Add(
                            new Grid_element((grid_alphabet[i].word + grid_alphabet[j].word).Replace(" ", ""),
                        grid_alphabet[i].symbol + grid_alphabet[j].symbol));

                        ind++;
                    }
                }

                grids[0].count = grids[0].elements.Count;
                grid_ind = 0;
            }

            if(MW != null)
            {
                //need to close mousegrid window or there will be 1 more window every time you change
                //mousegrid settings that require mousegrid regeneration
                //you can easily see this windows by pressing  windows key + tab
                MW.Close();
            }
            MW = new MouseGrid(grid_width, grid_height, grid_lines, grid_type, font_family, font_size, color_bg,
                color_font, rows_nr, cols_nr, figure_width, figure_height, grids[0].elements);
            //MW.regenerate_grid_symbols();

            grid_list = new List<string>();

            for (int i = 0; i < cancels_str.Length; i++)
            {
                grid_list.Add(cancels_str[i]);
            }

            for (int i = 0; i < directions_str.Length; i++)
            {
                grid_list.Add(directions_str[i]);
            }

            for (int i = 0; i < drag_edges_str.Length; i++)
            {
                grid_list.Add(drag_edges_str[i]);
            }

            for (int i = 0; i < count; i++)
            {
                grid_list.Add(grid_alphabet[i].word);
            }

            for (int i = 0; i < count; i++)
            {
                for (int j = 0; j < count; j++)
                {
                    grid_list.Add(grid_alphabet[i].word + " " + grid_alphabet[j].word);
                }
            }

            for (int i = 0; i < count; i++)
            {
                grid_list.Add(grid_alphabet[i].word + " twice");
            }

            int r1, r2;
            for (int i = 0; i < count; i++)
            {
                for (int j = 0; j < count; j++)
                {
                    if (int.TryParse(grid_alphabet[i].symbol, out r1)
                        && int.TryParse(grid_alphabet[j].symbol, out r2))
                    {
                        grid_list.Add(grid_alphabet[i].symbol + grid_alphabet[j].symbol);
                    }
                }
            }

            return grid_list;
        }

        List<string> create_builtin_commands_list()
        {
            List<string> commands_list = new List<string>();

            if (are_all_bic_general_and_mouse_disabled() == false)
            {
                List<string> multi = new List<string>();
                multi.Add("twice");

                for (int i = 0; i < list_bic_general_and_mouse.Count(); i++)
                {
                    if (list_bic_general_and_mouse[i].enabled
                        && list_bic_general_and_mouse[i].type != bic_type.move_up
                        && list_bic_general_and_mouse[i].type != bic_type.move_down
                        && list_bic_general_and_mouse[i].type != bic_type.move_left
                        && list_bic_general_and_mouse[i].type != bic_type.move_right
                        && list_bic_general_and_mouse[i].type != bic_type.open_app
                        && list_bic_general_and_mouse[i].type != bic_type.switch_to_app)
                    {
                        commands_list.Add(list_bic_general_and_mouse[i].name);

                        if (list_bic_general_and_mouse[i].key_combination == "Yes")
                        {
                            for (int j = 0; j < s_combo.Length; j++)
                            {
                                commands_list.Add(s_combo[j] + " " + list_bic_general_and_mouse[i].name);
                            }

                            for (int j = 0; j < s_combo.Length; j++)
                            {
                                for (int k = 0; k < s_combo2.Length; k++)
                                {
                                    commands_list.Add(s_combo[j] + " " + s_combo2[k] + " " + list_bic_general_and_mouse[i].name);
                                }
                            }
                        }

                        if (list_bic_general_and_mouse[i].max_executions > 1)
                        {
                            multi = new List<string>();
                            multi.Add("twice");

                            for (int j = 2; j <= list_bic_general_and_mouse[i].max_executions; j++)
                            {
                                multi.Add(j + " times");
                            }

                            for (int k = 0; k < multi.Count; k++)
                            {
                                commands_list.Add(list_bic_general_and_mouse[i].name + " " + multi[k]);
                            }
                        }

                        if (list_bic_general_and_mouse[i].key_combination == "Yes"
                            && list_bic_general_and_mouse[i].max_executions > 1)
                        {
                            for (int j = 0; j < s_combo.Length; j++)
                            {
                                for (int k = 0; k < multi.Count; k++)
                                {
                                    commands_list.Add(s_combo[j] + " " + list_bic_general_and_mouse[i].name + " " + multi[k]);
                                }
                            }

                            for (int j = 0; j < s_combo.Length; j++)
                            {
                                for (int k = 0; k < s_combo2.Length; k++)
                                {
                                    for (int m = 0; m < multi.Count; m++)
                                    {
                                        commands_list.Add(s_combo[j] + " " + s_combo2[k] + " " + list_bic_general_and_mouse[i].name + " " + multi[m]);
                                    }
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < list_bic_keys_pressing.Count(); i++)
                {
                    if (list_bic_keys_pressing[i].enabled)
                    {
                        commands_list.Add(list_bic_keys_pressing[i].name);

                        if (list_bic_keys_pressing[i].key_combination == "Yes")
                        {
                            for (int j = 0; j < s_combo.Length; j++)
                            {
                                commands_list.Add(s_combo[j] + " " + list_bic_keys_pressing[i].name);
                            }

                            for (int j = 0; j < s_combo.Length; j++)
                            {
                                for (int k = 0; k < s_combo2.Length; k++)
                                {
                                    commands_list.Add(s_combo[j] + " " + s_combo2[k] + " " + list_bic_keys_pressing[i].name);
                                }
                            }
                        }

                        if (list_bic_keys_pressing[i].max_executions > 1)
                        {
                            multi = new List<string>();
                            multi.Add("twice");

                            for (int j = 2; j <= list_bic_keys_pressing[i].max_executions; j++)
                            {
                                multi.Add(j + " times");
                            }

                            for (int k = 0; k < multi.Count; k++)
                            {
                                commands_list.Add(list_bic_keys_pressing[i].name + " " + multi[k]);
                            }
                        }

                        if (list_bic_keys_pressing[i].key_combination == "Yes"
                            && list_bic_keys_pressing[i].max_executions > 1)
                        {
                            for (int j = 0; j < s_combo.Length; j++)
                            {
                                for (int k = 0; k < multi.Count; k++)
                                {
                                    commands_list.Add(s_combo[j] + " " + list_bic_keys_pressing[i].name + " " + multi[k]);
                                }
                            }

                            for (int j = 0; j < s_combo.Length; j++)
                            {
                                for (int k = 0; k < s_combo2.Length; k++)
                                {
                                    for (int m = 0; m < multi.Count; m++)
                                    {
                                        commands_list.Add(s_combo[j] + " " + s_combo2[k] + " " + list_bic_keys_pressing[i].name + " " + multi[m]);
                                    }
                                }
                            }
                        }

                        /* This is the right solution, but it overloads the dictionary
                         * (left 100 isn't recognized)
                         * Compile time increased by 7 seconds when this is used
                        if (list_bic_keys_pressing[i].key_combination == "Yes" && list_bic_keys_pressing[i].max_executions > 1)
                        {
                            for (int m = 0; m < list_bic_keys_pressing.Count(); m++)
                            {
                                if (list_bic_keys_pressing[m].key_combination == "Yes" && list_bic_keys_pressing[m].max_executions > 1)
                                {
                                    multi = new Choices(new string[] { "twice" });

                                    for (int j = 2; j <= list_bic_keys_pressing[i].max_executions
                                                 && j <= list_bic_keys_pressing[m].max_executions; j++)
                                    {
                                        multi.Add(new string[] { j + " times" });
                                    }

                                    gb = new GrammarBuilder();
                                    gb.Append(list_bic_keys_pressing[m].name);
                                    gb.Append(list_bic_keys_pressing[i].name);
                                    gb.Append(multi);
                                    ch.Add(gb);

                                    gb = new GrammarBuilder();
                                    gb.Append(combo);
                                    gb.Append(list_bic_keys_pressing[m].name);
                                    gb.Append(list_bic_keys_pressing[i].name);
                                    gb.Append(multi);
                                    ch.Add(gb);
                                }
                            }
                        }
                        */
                    }
                }

                for (int i = 0; i < list_bic_char_inserting.Count(); i++)
                {
                    if (list_bic_char_inserting[i].enabled)
                    {
                        commands_list.Add(list_bic_char_inserting[i].name);

                        if (list_bic_char_inserting[i].max_executions > 1)
                        {
                            multi = new List<string>();
                            multi.Add("twice");

                            for (int j = 2; j <= list_bic_char_inserting[i].max_executions; j++)
                            {
                                multi.Add(j + " times");
                            }

                            for (int k = 0; k < multi.Count; k++)
                            {
                                commands_list.Add(list_bic_char_inserting[i].name + " " + multi[k]);
                            }
                        }
                    }
                }

                List<int> pixels = new List<int>();

                for (int i = 1; i <= 100; i++)
                {
                    pixels.Add(i);
                }

                List<string> enabled_mouse_moves = new List<string>();

                if (is_bic_in_general_and_mouse_enabled(bic_type.move_up))
                    enabled_mouse_moves.Add(s_mouse_moves[0]);
                if (is_bic_in_general_and_mouse_enabled(bic_type.move_down))
                    enabled_mouse_moves.Add(s_mouse_moves[1]);
                if (is_bic_in_general_and_mouse_enabled(bic_type.move_left))
                    enabled_mouse_moves.Add(s_mouse_moves[2]);
                if (is_bic_in_general_and_mouse_enabled(bic_type.move_right))
                    enabled_mouse_moves.Add(s_mouse_moves[3]);

                if (enabled_mouse_moves.Count > 0)
                {
                    for (int i = 0; i < enabled_mouse_moves.Count; i++)
                    {
                        for (int j = 0; j < pixels.Count; j++)
                        {
                            commands_list.Add(enabled_mouse_moves[i] + " " + pixels[j]);
                        }
                    }
                }
            }

            return commands_list;
        }
                
        //returns custom command actions for recognized string (returns null if not found)
        List<CC_Action> get_custom_command_actions(string str)
        {
            //custom commands for specific program have higher priority than for any program

            str = str.ToLower();
            string name;

            foreach (string program in new string[] 
                { 
                    last_foreground_window_process.ToLower(), Middle_Man.any_program_name.ToLower() 
                })
            {
                foreach (Profile p in Middle_Man.profiles)
                {
                    string profile_program = p.program.ToLower();

                    if (profile_program == program && p.enabled)
                    {
                        foreach (CustomCommand cc in p.custom_commands)
                        {
                            name = cc.name.ToLower();

                            if (cc.enabled)
                            {
                                if (name == str)
                                    return cc.actions;

                                if (cc.max_executions >= 2 && str == name + " twice")
                                    return cc.actions;

                                for (int i = 2; i <= cc.max_executions; i++)
                                {
                                    if (str == name + " " + i + " times")
                                        return cc.actions;
                                }
                            }
                        }
                    }
                }
            }

            return null;
        }

        List<string> create_custom_commands_list(string program)
        {
            List<string> ccommands_list = new List<string>();
            List<string> multi = new List<string>(); ;
            multi.Add("twice");

            program = program.ToLower();

            int cc2_commands_found = 0;

            List<CustomCommand> unique_commands = new List<CustomCommand>();

            foreach (Profile p in Middle_Man.profiles)
            {
                if(p.program.ToLower() == program && p.enabled)
                {
                    foreach(CustomCommand cc in p.custom_commands)
                    {
                        if (cc.enabled)
                        {
                            bool found_name = false;
                            bool found_both = false;

                            foreach (CustomCommand unique in unique_commands)
                            {
                                if (unique.name == cc.name)
                                {
                                    found_name = true;

                                    if (unique.max_executions == cc.max_executions)
                                    {
                                        found_both = true;

                                        break;
                                    }
                                }
                            }

                            if (found_name == false)
                            {
                                ccommands_list.Add(cc.name);

                                CustomCommand u = new CustomCommand();
                                u.name = cc.name;
                                u.max_executions = cc.max_executions;

                                unique_commands.Add(u);
                            }

                            if (found_both == false && cc.max_executions > 1)
                            {
                                multi = new List<string>();
                                multi.Add("twice");

                                for (int j = 2; j <= cc.max_executions; j++)
                                {
                                    multi.Add(j + " times");
                                }

                                for (int j = 0; j < multi.Count; j++)
                                {
                                    ccommands_list.Add(cc.name + " " + multi[j]);
                                }

                                CustomCommand u = new CustomCommand();
                                u.name = cc.name;
                                u.max_executions = cc.max_executions;

                                unique_commands.Add(u);
                            }

                            cc2_commands_found++;
                        }
                    }
                }
            }

            return ccommands_list;
        }

        void adv_mouse()
        {
            recognition_suspended = true;

            if (smart_grid)
                Thread.Sleep(50);//min=25 here and min=10 in mouse.cs

            bool f = false, cancel = false;

            if (cancels_str.Contains(r) == true)
                cancel = true;

            if (cancel == false)
            {
                int x = 0, y = 0;

                //if (last_command == bic_type.drop && (r == drag_edges_str[0] 
                if ((r == drag_edges_str[0].Replace(" ", "") || r == drag_edges_str[1].Replace(" ", "")
                    || r == drag_edges_str[2].Replace(" ", "") || r == drag_edges_str[3].Replace(" ", "")
                    || r == drag_edges_str[4].Replace(" ", "")))
                {
                    f = true;

                    if (last_command != bic_type.drop)
                    {
                        if (control)
                            key_down(VirtualKeyCode.CONTROL);
                        if (shift)
                            key_down(VirtualKeyCode.SHIFT);
                        if (alt)
                            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
                        if (windows)
                            key_down(VirtualKeyCode.LWIN);
                    }

                    if (r == drag_edges_str[0].Replace(" ", ""))
                    {
                        x = (int)(screen_width / 2);
                        y = 0;
                    }
                    else if (r == drag_edges_str[1].Replace(" ", ""))
                    {
                        x = (int)(screen_width / 2);
                        y = screen_height - 1;
                    }
                    else if (r == drag_edges_str[2].Replace(" ", ""))
                    {
                        y = (int)(screen_height / 2);
                        x = 0;
                    }
                    else if (r == drag_edges_str[3].Replace(" ", ""))
                    {
                        y = (int)(screen_height / 2);
                        x = screen_width - 1;
                    }
                    else if (r == drag_edges_str[4].Replace(" ", ""))
                    {
                        y = (int)(screen_height / 2);
                        x = (int)(screen_width / 2);
                    }

                    if (last_command == bic_type.move)
                    {
                        real_mouse_move(x, y);
                    }
                    else if (last_command == bic_type.left)
                    {
                        LMBClick(x, y);
                    }
                    else if (last_command == bic_type.right)
                    {
                        RMBClick(x, y);
                    }
                    else if (last_command == bic_type.double2)
                    {
                        DLMBClick(x, y);
                    }
                    else if (last_command == bic_type.triple)
                    {
                        TLMBClick(x, y);
                    }
                    else if (last_command == bic_type.drop)
                    {
                        freeze_mouse(50);
                        real_mouse_move(x, y);
                        freeze_mouse(50);
                        left_up();
                    }
                    else if (last_command == bic_type.drag)
                    {
                        LMBHold(x, y);

                        click_times++;

                        if (read_recognized_speech) ss.SpeakAsync("drop");
                    }

                    click_times--;

                    if (control)
                        key_up(VirtualKeyCode.CONTROL);
                    if (shift)
                        key_up(VirtualKeyCode.SHIFT);
                    if (alt)
                        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                    if (windows)
                        key_up(VirtualKeyCode.LWIN);
                }
                else
                    r = r.Replace(" ", "");

                if (r.Length > 5 && r.Substring(r.Length - 5) == "twice")
                {
                    r = r.Remove(r.Length - 5);
                    r = r + r;
                }

                for (int i = 0; i < grids[grid_ind].elements.Count && f == false; i++)
                {
                    if (r == grids[grid_ind].elements[i].word
                        || r == grids[grid_ind].elements[i].symbol)
                    {
                        f = true;

                        if (smart_grid)
                        {
                            int ind = -1;
                            int length = grids[grid_ind].elements[i].word_length;

                            for (int j = 0; j < grids[grid_ind].elements.Count; j++)
                            {
                                if (grids[grid_ind].elements[j].word_length < length
                                    && grids[grid_ind].elements[j].count
                                    <= grids[grid_ind].elements[i].count)
                                {
                                    ind = j;
                                    length = grids[grid_ind].elements[j].word_length;
                                }
                            }

                            if (ind != -1)
                            {
                                string symbol = grids[grid_ind].elements[ind].symbol;
                                string word = grids[grid_ind].elements[ind].word;
                                uint count = grids[grid_ind].elements[ind].count;

                                grids[grid_ind].elements[ind].count = grids[grid_ind].elements[i].count;
                                grids[grid_ind].elements[ind].symbol = grids[grid_ind].elements[i].symbol;
                                grids[grid_ind].elements[ind].word = grids[grid_ind].elements[i].word;
                                grids[grid_ind].elements[ind].word_length
                                    = grids[grid_ind].elements[i].word_length;

                                grids[grid_ind].elements[i].count = count;
                                grids[grid_ind].elements[i].symbol = symbol;
                                grids[grid_ind].elements[i].word = word;
                                grids[grid_ind].elements[i].word_length = length;
                            }

                            grids[grid_ind].elements[i].count++;
                        }

                        //delayed Hide, because Mousegrid must be shown when figures content is updated
                        //or it would update when shown next time (figure content change would be noticed
                        //by user)
                        if (smart_grid)
                        {
                            MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                MW.Hide();
                            }));
                        }

                        if (grid_type == GridType.hexagonal)
                        {
                            int two_rows = (int)(i / (cols_nr * 2 - 1));
                            int m = i % (cols_nr * 2 - 1);

                            if (m < cols_nr)//first
                            {
                                x = (int)Math.Round((double)(screen_width * (m + 0.5) / cols_nr));
                                y = (int)Math.Round((double)(screen_height * (two_rows + 0.5) / rows_nr));
                            }
                            else
                            {
                                x = (int)Math.Round((double)(screen_width * (m - cols_nr + 1) / cols_nr));
                                y = (int)Math.Round((double)(screen_height * (two_rows + 1) / rows_nr));
                            }
                        }
                        else
                        {
                            double offest_v = 0.5;
                            double offest_h = 0.5;

                            if (grid_type == MainWindow.GridType.square_horizontal_precision
                                    || grid_type == MainWindow.GridType.square_combined_precision)
                            {
                                int col_nr = i % cols_nr;
                                if (col_nr % 3 == 0)
                                    offest_v = 0.25;
                                else if (col_nr % 3 == 1)
                                    offest_v = 0.5;
                                else
                                    offest_v = 0.75;
                            }
                            if (grid_type == MainWindow.GridType.square_vertical_precision
                                || grid_type == MainWindow.GridType.square_combined_precision)
                            {
                                int row_nr = (int)(i / cols_nr);
                                if (row_nr % 3 == 0)
                                    offest_h = 0.25;
                                else if (row_nr % 3 == 1)
                                    offest_h = 0.5;
                                else
                                    offest_h = 0.75;
                            }

                            x = (int)Math.Round((double)(screen_width *
                                (i % cols_nr + offest_h) / cols_nr));
                            y = (int)Math.Round((double)(screen_height *
                                ((int)(i / cols_nr) + offest_v) / rows_nr));
                        }

                        x += offset_x;
                        y += offset_y;

                        if (x < 0) x = 0;
                        else if (x > screen_width - 1) x = screen_width - 1;

                        if (y < 0) y = 0;
                        else if (y > screen_height - 1) y = screen_height - 1;

                        if (control)
                            key_down(VirtualKeyCode.CONTROL);
                        if (shift)
                            key_down(VirtualKeyCode.SHIFT);
                        if (alt)
                            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
                        if (windows)
                            key_down(VirtualKeyCode.LWIN);

                        if (last_command == bic_type.move)
                        {
                            real_mouse_move(x, y);
                        }
                        else if (last_command == bic_type.left)
                        {
                            LMBClick(x, y);
                        }
                        else if (last_command == bic_type.right)
                        {
                            RMBClick(x, y);
                        }
                        else if (last_command == bic_type.double2)
                        {
                            DLMBClick(x, y);
                        }
                        else if (last_command == bic_type.triple)
                        {
                            TLMBClick(x, y);
                        }
                        else if (last_command == bic_type.drop)
                        {
                            freeze_mouse(50);
                            real_mouse_move(x, y);
                            freeze_mouse(50);
                            left_up();
                        }
                        else if (last_command == bic_type.drag)
                        {
                            LMBHold(x, y);

                            click_times++;

                            if (read_recognized_speech) ss.SpeakAsync("drop");
                        }

                        click_times--;

                        if (last_command != bic_type.drag)
                        {
                            if (control)
                                key_up(VirtualKeyCode.CONTROL);
                            if (shift)
                                key_up(VirtualKeyCode.SHIFT);
                            if (alt)
                                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                            if (windows)
                                key_up(VirtualKeyCode.LWIN);
                        }
                    }
                }
            }
            else if (f == false)
            {
                if (last_command == bic_type.drop)
                    left_up();

                last_command = bic_type.cancel;
            }

            if ((click_times > 0 && last_command != bic_type.cancel)
                || last_command == bic_type.drag)
            {
                enable_grid();
                if (smart_grid)
                {
                    MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                    {
                        MW.Hide(); //for TopMost
                        MW.Opacity = 1;
                        MW.Show();
                    }));
                }
                else
                {
                    MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                    {
                        MW.Show();
                    }));
                }
                grid_visible = true;
            }
            else
            {
                SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                {
                    SW.Bmode.Visibility = Visibility.Visible;
                }));

                current_mode = mode.command;
                load_turned_on();
            }

            if (last_command == bic_type.drag)
            {
                last_command = bic_type.drop;
            }
            else if (last_command == bic_type.drop && click_times > 0)
            {
                last_command = bic_type.drag;
            }

            if (movable_grid)
            {
                offset_x = offset_y = 0;

                MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                {
                    MW.Top = offset_y;
                    MW.Left = offset_x;
                }));
            }

            recognition_suspended = false;
        }

        int click_times;

        bool control = false, shift = false, alt = false, windows = false;

        string r, r_lowercase; //r = recognized speech
        bool speech_recognized = false;

        private readonly object lock_list_cc_any = new object();
        private readonly object lock_list_cc_foreground = new object();
        private readonly object lock_list_switch_to_apps = new object();

        // Handle the Speech Recognized event
        void WaveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            try
            {
                if (recognition_suspended == false)
                {
                    inside_speech_recognized_event = true;

                    if (recognizer == null)
                        return;

                    if (recognizer.AcceptWaveform(e.Buffer, e.BytesRecorded) && recognition_suspended == false)
                    {
                        string json = recognizer.Result();

                        JsonDocument doc = JsonDocument.Parse(json);

                        r = doc.RootElement.GetProperty("text").GetString() ?? ""; //r = recognized speech
                        
                        int c = 0; //speech recognition confidence

                        if (r != null && r != "")
                        {
                            speech_recognized = true;

                            if (ss.Volume != ss_volume)
                                ss.Volume = ss_volume;

                            if (current_mode == mode.off)
                            {
                                int ind = -1; //index of highest confidence word
                                int c_curr;

                                for (int i = 0; i < list_off_mode.Count; i++)
                                {
                                    c_curr = (int)get_similarity(r, list_off_mode[i]);

                                    if (c_curr > c)
                                    {
                                        c = c_curr;
                                        ind = i;
                                    }
                                }

                                if (ind != -1)
                                    r = list_off_mode[ind];
                                else
                                    r = "";

                                if (c >= confidence_turning_on)
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);

                                    current_mode = last_mode;

                                    load_turned_on(true);

                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        if (r.Length > 0)
                                            SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                        SW.TBconfidence.Text = c.ToString() + "/" + confidence_turning_on;

                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                            = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                    }));
                                }
                                else
                                {
                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        if (r.Length > 0)
                                            SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                        SW.TBconfidence.Text = c.ToString() + "/" + confidence_turning_on;

                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                            = new SolidColorBrush(Color.FromRgb(232, 4, 4));
                                    }));
                                }
                            }
                            else if (current_mode == mode.command)
                            {
                                r = replace_number_words(r);

                                r = r.Replace("juliet", "juliett");
                                r = r.Replace("alpha", "alfa");
                                r = r.Replace("pays", "paste");
                                r = r.Replace("based", "paste");
                                r = r.Replace("council", "cancel");
                                r = r.Replace("back space", "backspace");
                                r = r.Replace("tax space", "backspace");
                                r = r.Replace("max base", "backspace");
                                r = r.Replace("tax base", "backspace");
                                r = r.Replace("killer", "kilo");
                                r = r.Replace("lemme", "lima");
                                r = r.Replace("column", "colon");
                                r = r.Replace("colin", "colon");
                                r = r.Replace("cullen", "colon");
                                r = r.Replace("lb", "pound");
                                r = r.Replace("carrot", "caret");
                                r = r.Replace("clegg", "click");
                                r = r.Replace("tremble", "triple");
                                r = r.Replace("treble", "triple");
                                r = r.Replace("tribble", "triple");
                                r = r.Replace("evil", "ouble");
                                r = r.Replace("oh then", "open");
                                r = r.Replace("friend", "print");
                                r = r.Replace("chap", "tab");
                                r = r.Replace("you tube", "new tab");
                                r = r.Replace("brad shaw", "bravo");
                                r = r.Replace("fox trot", "foxtrot");
                                r = r.Replace("x ray", "xray");
                                r = r.Replace("hi fen", "hyphen");
                                r = r.Replace("hi finn", "hyphen");
                                r = r.Replace("carried", "carot");
                                r = r.Replace("that quote", "back quote");
                                r = r.Replace("amber's and", "ampersand");
                                r = r.Replace("amber's end", "ampersand");
                                r = r.Replace("am percent", "ampersand");
                                r = r.Replace("sent", "cent");
                                r = r.Replace("drug", "drag");
                                r = r.Replace("offer", "alpha");
                                r = r.Replace("alva", "alpha");
                                r = r.Replace("cold", "hold");
                                r = r.Replace("injured", "enter");
                                r = r.Replace("answered", "enter");
                                r = r.Replace("weiss", "twice");
                                r = r.Replace("wise", "twice");
                                r = r.Replace("pains", "paint");
                                r = r.Replace("out", "alt");
                                r = r.Replace("added", "alt");
                                r = r.Replace("act three", "xray");
                                r = r.Replace("hum", "home");
                                r = r.Replace("palm", "home");
                                r = r.Replace("pay job", "page up");
                                r = r.Replace("algo", "echo");
                                r = r.Replace("glove", "golf");
                                r = r.Replace("julian", "juliett");
                                r = r.Replace("good back", "quebec");
                                r = r.Replace("axe three", "xray");
                                r = r.Replace("bug", "back");
                                r = r.Replace("buck", "back");
                                r = r.Replace("hush", "hash");
                                r = r.Replace("nov twice", "november twice");
                                r = r.Replace("tumblr", "tab");
                                r = r.Replace("tampa", "tab");
                                r = r.Replace("thumper", "tab");
                                r = r.Replace("andrew", "undo");
                                r = r.Replace("i do", "undo");
                                r = r.Replace("calmer", "comma");
                                r = r.Replace("come on", "comma");
                                r = r.Replace("tacoma", "comma");
                                r = r.Replace("rhythm", "redo");
                                r = r.Replace("riddle", "redo");
                                r = r.Replace("read though", "redo");
                                //r = r.Replace("", "");

                                int ind1 = 0, ind2 = 0, ind3 = 0, ind4 = 0, ind5 = 0; //indexes of highest confidence words
                                int c1 = 0, c2 = 0, c3 = 0, c4 = 0, c5 = 0; //confidences for each list
                                int c_curr;
                                c = 0;

                                lock (lock_list_cc_foreground)
                                {
                                    for (int i = 0; i < list_cc_foreground.Count; i++)
                                    {
                                        c_curr = (int)get_similarity(r, list_cc_foreground[i]);

                                        if (c_curr > c1)
                                        {
                                            c1 = c_curr;
                                            ind1 = i;
                                        }
                                    }
                                }

                                lock (lock_list_cc_any)
                                {
                                    for (int i = 0; i < list_cc_any.Count; i++)
                                    {
                                        c_curr = (int)get_similarity(r, list_cc_any[i]);

                                        if (c_curr > c2)
                                        {
                                            c2 = c_curr;
                                            ind2 = i;
                                        }
                                    }
                                }

                                if (apps_opening)
                                {
                                    for (int i = 0; i < list_open_apps.Count; i++)
                                    {
                                        c_curr = (int)get_similarity(r, list_open_apps[i]);

                                        if (c_curr > c3)
                                        {
                                            c3 = c_curr;
                                            ind3 = i;
                                        }
                                    }
                                }

                                if (apps_switching)
                                {
                                    lock (lock_list_switch_to_apps)
                                    {
                                        for (int i = 0; i < list_switch_to_apps.Count; i++)
                                        {
                                            c_curr = (int)get_similarity(r, list_switch_to_apps[i]);

                                            if (c_curr > c4)
                                            {
                                                c4 = c_curr;
                                                ind4 = i;
                                            }
                                        }
                                    }
                                }

                                for (int i = 0; i < list_current.Count; i++)
                                {
                                    c_curr = (int)get_similarity(r, list_current[i]);

                                    if (c_curr > c5)
                                    {
                                        c5 = c_curr;
                                        ind5 = i;
                                    }
                                }

                                lock (lock_list_cc_foreground)
                                {
                                    if (list_cc_foreground.Count > 0)
                                    {
                                        if (c1 > c)
                                        {
                                            c = c1;
                                            r = list_cc_foreground[ind1];
                                        }
                                    }
                                }

                                lock (lock_list_cc_any)
                                {
                                    if (list_cc_any.Count > 0)
                                    {
                                        if (c2 > c)
                                        {
                                            c = c2;
                                            r = list_cc_any[ind2];
                                        }
                                    }
                                }

                                if (apps_opening && list_open_apps.Count > 0)
                                {
                                    if (c3 > c)
                                    {
                                        c = c3;
                                        r = list_open_apps[ind3];
                                    }
                                }

                                lock (lock_list_switch_to_apps)
                                {
                                    if (apps_switching && list_switch_to_apps.Count > 0)
                                    {
                                        if (c4 > c)
                                        {
                                            c = c4;
                                            r = list_switch_to_apps[ind4];
                                        }
                                    }
                                }

                                if (list_current.Count > 0)
                                {
                                    if (c5 > c)
                                    {
                                        c = c5;
                                        r = list_current[ind5];
                                    }
                                }

                                string r_lowercase = r.ToLower();

                                if (c > 0)
                                {
                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                        SW.TBconfidence.Text = c.ToString() + "/" + confidence_other_commands;
                                        if (c >= confidence_other_commands)
                                        {
                                            SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                                = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                        }
                                        else
                                        {
                                            SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                                = new SolidColorBrush(Color.FromRgb(232, 4, 4));
                                        }
                                    }));
                                }

                                //B.Content = r + " | " + c + " | " + e.Result.ReplacementWordUnits.Count;
                                //B.Content = r + " | " + c + " | " + sem;

                                if (r_lowercase == "open computer")
                                    r = r_lowercase = "open computer";
                                if (r_lowercase == "open task manager")
                                    r = r_lowercase = "open task manager";
                                if (r_lowercase == "open settings")
                                    r = r_lowercase = "open settings";
                                if (r_lowercase == "open power menu")
                                    r = r_lowercase = "open power menu";

                                List<CC_Action> actions = get_custom_command_actions(r);

                                //Custom Command
                                if (c >= confidence_other_commands && actions != null)
                                {
                                    recognition_suspended = true;

                                    THRcommands = new Thread(() => execute_custom_commands(actions));
                                    THRcommands.Start();
                                }
                                //turn off speech recognition
                                else if (r == turn_off && c >= confidence_other_commands && is_bic_in_general_and_mouse_enabled(bic_type.turn_off))
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);

                                    load_turned_off(true);

                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                        SW.TBconfidence.Text = c.ToString() + "/" + confidence_other_commands;
                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                            = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                    }));
                                }
                                //switch to dictation mode
                                else if (r == switch_to_dictation_mode && c >= confidence_other_commands)
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);

                                    current_mode = mode.dictation;

                                    load_turned_on(true);
                                }
                                //show speech recognition window
                                else if (r == show_speech_recognition && c >= confidence_other_commands)
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);
                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        if (SW.IsLoaded)
                                        {
                                            SW.Show();
                                        }
                                        else
                                        {
                                            SW = new SpeechWindow();
                                            SW.Topmost = true;
                                            SW.ShowInTaskbar = false;
                                            load_coords();
                                            SW.Show();
                                        }
                                    }));

                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        SW.mode = 1;
                                        SW.Bmode.Content = mode_command;
                                        SW.Bmode.Foreground = new SolidColorBrush(Color.FromRgb(0, 128, 0));

                                        // Create a BitmapSource  
                                        BitmapImage bitmap = new BitmapImage();
                                        bitmap.BeginInit();
                                        bitmap.UriSource = new Uri(icon_command);
                                        bitmap.EndInit();

                                        SW.Icon = bitmap;

                                        SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                        SW.TBconfidence.Text = c.ToString() + "/" + confidence_other_commands;
                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                            = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                    }));
                                }
                                //hide speech recognition window
                                else if (r == hide_speech_recognition && c >= confidence_other_commands)
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);
                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        if (SW.IsVisible)
                                        {
                                            SW.Hide();
                                        }
                                    }));
                                }
                                //execute a command
                                else if (c >= confidence_other_commands)
                                {
                                    //if (r.Length >= 6 && r.Substring(0, 6).ToLower() != "press ")
                                    //    r.Replace("press ", "");
                                    recognition_suspended = true;

                                    THRcommands = new Thread(new ThreadStart(execute_commands));
                                    THRcommands.Start();
                                }
                            }
                            else if (current_mode == mode.grid)
                            {
                                r = replace_number_words(r);

                                r = r.Replace("alpha", "alfa");
                                r = r.Replace("council", "cancel");
                                r = r.Replace("killer", "kilo");
                                r = r.Replace("lemme", "lima");
                                r = r.Replace("brad shaw", "bravo");
                                r = r.Replace("fox trot", "foxtrot");
                                r = r.Replace("x ray", "xray");
                                r = r.Replace("amber's and", "ampersand");
                                r = r.Replace("amber's end", "ampersand");
                                r = r.Replace("am percent", "ampersand");
                                r = r.Replace("drug", "drag");
                                r = r.Replace("offer", "alpha");
                                r = r.Replace("alva", "alpha");
                                r = r.Replace("weiss", "twice");
                                r = r.Replace("wise", "twice");
                                r = r.Replace("act three", "xray");
                                r = r.Replace("algo", "echo");
                                r = r.Replace("glove", "golf");
                                r = r.Replace("julian", "juliett");
                                r = r.Replace("good back", "quebec");
                                r = r.Replace("axe three", "xray");
                                r = r.Replace("lb", "pound");
                                r = r.Replace("juliet", "juliett");
                                r = r.Replace("nov twice", "november twice");
                                r = r.Replace("column", "colon");
                                r = r.Replace("colin", "colon");
                                r = r.Replace("cullen", "colon");

                                //List<string> list = new List<string>();

                                //for (int i = 0; i < grids[grid_ind].elements.Count; i++)
                                //{
                                //    if (grids[grid_ind].elements[i].word.Contains("4"))
                                //    {
                                //        list.Add(grids[grid_ind].elements[i].word);
                                //    }
                                //}

                                //int z = 5;

                                //for (int i = 0; i < list_mousegrid.Count; i++)
                                //{
                                //    if (list_mousegrid[i] == "4 alfa")
                                //    {
                                //        string s = list_mousegrid[i];
                                //        int x = 5;
                                //    }
                                //}

                                int ind = -1; //indexes of highest confidence words
                                int c_curr;

                                for (int i = 0; i < list_mousegrid.Count; i++)
                                {
                                    c_curr = (int)get_similarity(r, list_mousegrid[i]);

                                    if (c_curr > c)
                                    {
                                        c = c_curr;
                                        ind = i;
                                    }
                                }

                                if (ind != -1)
                                    r = list_mousegrid[ind];
                                else
                                    r = "";

                                SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    if (r.Length > 0)
                                        SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                    SW.TBconfidence.Text = c.ToString() + "/" + confidence_other_commands;

                                    if (c >= confidence_other_commands)
                                    {
                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                            = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                    }
                                    else
                                    {
                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                            = new SolidColorBrush(Color.FromRgb(232, 4, 4));
                                    }
                                }));

                                //move mouse grid
                                if (c >= confidence_other_commands && grid_visible && movable_grid
                                && (r == directions_str[0] || r == directions_str[1]
                                || r == directions_str[2] || r == directions_str[3]))
                                {
                                    if (r == directions_str[0])
                                    {
                                        offset_y -= (int)Math.Round(offset * figure_height);
                                    }
                                    else if (r == directions_str[1])
                                    {
                                        offset_y += (int)Math.Round(offset * figure_height);
                                    }
                                    else if (r == directions_str[2])
                                    {
                                        offset_x -= (int)Math.Round(offset * figure_width);
                                    }
                                    else if (r == directions_str[3])
                                    {
                                        offset_x += (int)Math.Round(offset * figure_width);
                                    }

                                    MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        MW.Top = offset_y;
                                        MW.Left = offset_x;
                                    }));

                                    if (read_recognized_speech) ss.SpeakAsync(r);
                                }
                                //perform a mousegrid action
                                else if (grid_visible && c >= confidence_other_commands)
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);

                                    MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        if (smart_grid)
                                            MW.Opacity = 0; //hide is delayed until mousegrid content changes
                                        else
                                            MW.Hide();
                                    }));
                                    grid_visible = false;

                                    r = r.Replace(" ", "");

                                    THRmouse = new Thread(new ThreadStart(adv_mouse));
                                    THRmouse.Start();
                                }
                            }
                            else if (current_mode == mode.dictation)
                            {
                                if (r == "back space")
                                    r = "backspace";
                                else if (r == "tax space")
                                    r = "backspace";
                                else if (r == "max space")
                                    r = "backspace";
                                else if (r == "tax base")
                                    r = "backspace";
                                else if (r == "max base")
                                    r = "backspace";
                                else if (r == "column")
                                    r = "colon";
                                else if (r == "colin")
                                    r = "colon";
                                else if (r == "cullen")
                                    r = "colon";
                                else if (r == "chap")
                                    r = "tab";
                                else if (r == "hi fen")
                                    r = "hyphen";
                                else if (r == "hi finn")
                                    r = "hyphen";
                                else if (r == "tumblr")
                                    r = "tab";
                                else if (r == "tampa")
                                    r = "tab";
                                else if (r == "double")
                                    r = "tab";
                                else if (r == "thumper")
                                    r = "tab";
                                else if (r == "andrew")
                                    r = "undo";
                                else if (r == "i do")
                                    r = "undo";
                                else if (r == "calmer")
                                    r = "comma";
                                else if (r == "come on")
                                    r = "comma";
                                else if (r == "tacoma")
                                    r = "comma";
                                else if (r == "rhythm")
                                    r = "redo";
                                else if (r == "riddle")
                                    r = "redo";
                                else if (r == "read though")
                                    r = "redo";
                                else if (r == "the red line")
                                    r = "delete line";

                                int ind1 = 0; //index of highest confidence words
                                int c1 = 0; //confidences for each list
                                int c_curr;

                                for (int i = 0; i < list_current.Count; i++)
                                {
                                    c_curr = (int)get_similarity(r, list_current[i]);

                                    if (c_curr > c1)
                                    {
                                        c1 = c_curr;
                                        ind1 = i;
                                    }
                                }

                                if (c1 >= confidence_other_commands)
                                {
                                    r = list_current[ind1];
                                    c = c1;
                                }

                                bool dictation_command = false;

                                if (c >= confidence_other_commands)
                                {
                                    foreach (BuiltInCommand bic in list_bic_dictation)
                                    {
                                        if (r == bic.name && bic.enabled)
                                        {
                                            dictation_command = true;
                                            break;
                                        }
                                    }
                                }

                                //turn off
                                if (r == turn_off && c >= confidence_other_commands && is_bic_in_dictation_enabled(bic_type.turn_off))
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);

                                    load_turned_off(true);

                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                        SW.TBconfidence.Text = c.ToString() + "/" + confidence_other_commands;
                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                            = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                    }));
                                }
                                //switch to command mode
                                else if ((int)get_similarity(r, switch_to_command_mode) >= c
                                    && (int)get_similarity(r, switch_to_command_mode) >= confidence_other_commands
                                    && is_bic_in_dictation_enabled(bic_type.switch_to_command))
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);

                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                        SW.TBconfidence.Text = ((int)get_similarity(r, switch_to_command_mode)).ToString() + "/" + confidence_other_commands;
                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                            = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                    }));

                                    current_mode = mode.command;

                                    load_turned_on(true);
                                }
                                else if(dictation_command)
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);

                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                                        SW.TBconfidence.Text = c + "/" + confidence_other_commands;
                                        SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                    }));

                                    if (r == "uppercase")
                                    {
                                        uppercase_sentence = true;
                                    }
                                    else if (r == "lowercase")
                                    {
                                        uppercase_sentence = false;
                                    }
                                    else if (r == "undo")
                                    {
                                        sim.Keyboard.KeyDown(VirtualKeyCode.LCONTROL);
                                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_Z);
                                        Thread.Sleep(50);
                                        sim.Keyboard.KeyUp(VirtualKeyCode.VK_Z);
                                        sim.Keyboard.KeyUp(VirtualKeyCode.LCONTROL);
                                    }
                                    else if (r == "redo")
                                    {
                                        sim.Keyboard.KeyDown(VirtualKeyCode.LCONTROL);
                                        sim.Keyboard.KeyDown(VirtualKeyCode.VK_Y);
                                        Thread.Sleep(50);
                                        sim.Keyboard.KeyUp(VirtualKeyCode.VK_Y);
                                        sim.Keyboard.KeyUp(VirtualKeyCode.LCONTROL);
                                    }
                                    else if (r == "delete word")
                                    {
                                        sim.Keyboard.KeyDown(VirtualKeyCode.LCONTROL);
                                        sim.Keyboard.KeyDown(VirtualKeyCode.BACK);
                                        Thread.Sleep(50);
                                        sim.Keyboard.KeyUp(VirtualKeyCode.BACK);
                                        sim.Keyboard.KeyUp(VirtualKeyCode.LCONTROL);
                                    }
                                    else if (r == "delete line")
                                    {
                                        sim.Keyboard.KeyPress(VirtualKeyCode.END);

                                        Thread.Sleep(50);

                                        sim.Keyboard.KeyDown(VirtualKeyCode.LSHIFT);
                                        sim.Keyboard.KeyDown(VirtualKeyCode.HOME);
                                        Thread.Sleep(50);
                                        sim.Keyboard.KeyUp(VirtualKeyCode.HOME);
                                        sim.Keyboard.KeyUp(VirtualKeyCode.LSHIFT);

                                        Thread.Sleep(50);

                                        sim.Keyboard.KeyPress(VirtualKeyCode.BACK);

                                        uppercase_sentence = true;
                                    }
                                    else if (r == "space")
                                    {
                                        sim.Keyboard.KeyPress(VirtualKeyCode.SPACE);
                                    }
                                    else if (r == "left")
                                    {
                                        sim.Keyboard.KeyPress(VirtualKeyCode.LEFT);
                                    }
                                    else if (r == "right")
                                    {
                                        sim.Keyboard.KeyPress(VirtualKeyCode.RIGHT);
                                    }
                                    else if (r == "enter")
                                    {
                                        sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                                    }
                                    else if (r == "tab")
                                    {
                                        sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
                                    }
                                    else if (r == "backspace")
                                    {
                                        sim.Keyboard.KeyPress(VirtualKeyCode.BACK);
                                    }
                                    else if (r == "comma")
                                    {
                                        sim.Keyboard.TextEntry(", ");
                                        uppercase_sentence = false;
                                    }
                                    else if (r == "dot")
                                    {
                                        sim.Keyboard.TextEntry(". ");
                                        uppercase_sentence = true;
                                    }
                                    else if (r == "period")
                                    {
                                        sim.Keyboard.TextEntry(". ");
                                        uppercase_sentence = true;
                                    }
                                    else if (r == "hyphen")
                                    {
                                        sim.Keyboard.TextEntry("-");
                                        uppercase_sentence = false;
                                    }
                                    else if (r == "semicolon")
                                    {
                                        sim.Keyboard.TextEntry("; ");
                                        uppercase_sentence = false;
                                    }
                                    else if (r == "colon")
                                    {
                                        sim.Keyboard.TextEntry(": ");
                                        uppercase_sentence = false;
                                    }
                                    else if (r == "double quote")
                                    {
                                        sim.Keyboard.TextEntry("\"");
                                        uppercase_sentence = false;
                                    }
                                    else if (r == "quote")
                                    {
                                        sim.Keyboard.TextEntry("'");
                                        uppercase_sentence = false;
                                    }
                                    else if (r == "exclamation")
                                    {
                                        sim.Keyboard.TextEntry("! ");
                                        uppercase_sentence = true;
                                    }
                                    else if (r == "question")
                                    {
                                        sim.Keyboard.TextEntry("? ");
                                        uppercase_sentence = true;
                                    }
                                    else if (r == "open parenthesis")
                                    {
                                        sim.Keyboard.TextEntry("(");
                                        uppercase_sentence = false;
                                    }
                                    else if (r == "close parenthesis")
                                    {
                                        sim.Keyboard.TextEntry(")");
                                    }
                                }
                                //dictation text
                                else
                                //&& (r.Length < 6 || r.Substring(0, 6).ToLower() != "press "))
                                {
                                    if (read_recognized_speech) ss.SpeakAsync(r);

                                    if (uppercase_sentence)
                                    {
                                        r = r.FirstCharToUpper();
                                        uppercase_sentence = false;
                                    }

                                    sim.Keyboard.TextEntry(r);
                                    c = 100;

                                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        if (SW.IsVisible)
                                        {
                                            SW.TBrecognized_speech.Text = r;
                                            SW.TBconfidence.Text = c.ToString();
                                            SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                                = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                                        }
                                    }));
                                }
                            }
                        }
                    }

                    speech_recognized = false;
                    inside_speech_recognized_event = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW009", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        bool uppercase_sentence = true;

        void execute_commands()
        {
            recognition_suspended = true;

            try
            {
                if (read_recognized_speech)
                {
                    if(ss.Volume != ss_volume)
                        ss.Volume = ss_volume;

                    ss.SpeakAsync(r);
                }

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

                bool found = false;
                
                for (int i = 0; i < list_bic_keys_pressing.Count() && found == false; i++)
                {
                    if (r == list_bic_keys_pressing[i].name
                        && list_bic_keys_pressing[i].keys != null)
                    {
                        found = true;
                        r = list_bic_keys_pressing[i].keys;
                        a = r.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                        n = a.Length;
                    }
                }

                control = shift = alt = windows = false;

                a = r.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                n = a.Length;

                //mouse actions can use up to 2 //ctrl/shift/alt/win combo keys:
                if (a.Length == 3 &&
                    (a[2] == "move" || a[2] == "west" || a[2] == "east" || a[2] == "Double"
                     || a[2] == "triple" || a[2] == "drag" || a[2] == "click" || a[2] == "hold"
                     || a[2] == "scroll"))

                {
                    for (int k = 0; n > 1 && k <= 1; k++)
                    {
                        if (s_combo2.Contains(a[k]))
                        {
                            if (a[k] == s_combo2[0])
                            {
                                control = true;
                                r = r.Remove(0, s_combo2[0].Length);
                            }
                            else if (a[k] == s_combo2[1])
                            {
                                shift = true;
                                r = r.Remove(0, s_combo2[1].Length);
                            }
                            else if (a[k] == s_combo2[2]) //right alt would be recognized here
                            {
                                alt = true;
                                r = r.Remove(0, s_combo2[2].Length);
                            }
                            else if (a[k] == s_combo2[3])
                            {
                                windows = true;
                                r = r.Remove(0, s_combo2[3].Length);
                            }
                            r = r.Trim();
                        }
                    }
                }
                else
                {
                    if (n > 1) //ctrl/shift/alt/win combo detected
                    {
                        if (s_combo2.Contains(a[0]))
                        {
                            if (a[0] == s_combo2[0])
                            {
                                control = true;
                                r = r.Remove(0, s_combo2[0].Length);
                            }
                            else if (a[0] == s_combo2[1])
                            {
                                shift = true;
                                r = r.Remove(0, s_combo2[1].Length);
                            }
                            else if (a[0] == s_combo2[2]) //right alt shouldn't be recognized here
                            {
                                alt = true;
                                r = r.Remove(0, s_combo2[2].Length);
                            }
                            else if (a[0] == s_combo2[3])
                            {
                                windows = true;
                                r = r.Remove(0, s_combo2[3].Length);
                            }
                            r = r.Trim();
                        }
                    }
                }

                for (int j = 1; j <= executions; j++)
                {
                    if (control)
                        key_down(VirtualKeyCode.CONTROL);
                    if (shift)
                        key_down(VirtualKeyCode.SHIFT);
                    if (alt)
                        keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
                    if (windows)
                        key_down(VirtualKeyCode.LWIN);

                    found = false;

                    for (int i = 0; i < list_bic_general_and_mouse.Count() && found == false; i++)
                    {
                        if (list_bic_general_and_mouse[i].keys == null
                            && list_bic_general_and_mouse[i].enabled
                            && (r == list_bic_general_and_mouse[i].name
                            || (r.Contains(list_bic_general_and_mouse[i].name)
                            && r.Length > list_bic_general_and_mouse[i].name.Length
                            && list_bic_general_and_mouse[i].use_contains)))
                        {
                            found = true;

                            //debug only;
                            //MessageBox.Show(list_bic_general_and_mouse[i].name);

                            if (list_bic_general_and_mouse[i].type == bic_type.left
                                || list_bic_general_and_mouse[i].type == bic_type.right
                                || list_bic_general_and_mouse[i].type == bic_type.double2
                                || list_bic_general_and_mouse[i].type == bic_type.triple
                                || list_bic_general_and_mouse[i].type == bic_type.move
                                || list_bic_general_and_mouse[i].type == bic_type.drag)
                            {
                                if (smart_grid)
                                {
                                    IntPtr handle = GetForegroundWindow();
                                    string process_name = "";
                                    grid_ind = -1;

                                    Process[] arr = Process.GetProcesses();

                                    foreach (Process p in arr)
                                    {
                                        if (p.MainWindowHandle == handle)
                                        {
                                            process_name = p.ProcessName;
                                            break;
                                        }
                                    }

                                    for (int k = 1; k < grids.Count && grid_ind == -1; k++)
                                    {
                                        if (grids[k].process_name == process_name)
                                            grid_ind = k;
                                    }

                                    if (grid_ind == -1)
                                    {
                                        grids.Add(new Process_grid(process_name));
                                        grid_ind = grids.Count - 1;

                                        for (int k = 0; k < grids[0].elements.Count; k++)
                                        {
                                            grids[grid_ind].elements.Add(new Grid_element(
                                                grids[0].elements[k].word, grids[0].elements[k].symbol));
                                        }

                                        grids[grid_ind].count = grids[grid_ind].elements.Count;
                                    }

                                    //if(debug_mode) start_time();
                                    MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        MW.elements = grids[grid_ind].elements;
                                        MW.regenerate_grid_symbols();//1-211 ms
                                    }));
                                    //if(debug_mode) stop_time();
                                }
                                else grid_ind = 0;

                                last_command = list_bic_general_and_mouse[i].type;
                                enable_grid();
                                //if(debug_mode) start_time();
                                if (smart_grid)
                                {
                                    MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        MW.Hide();//for Topmost
                                        MW.Opacity = 1;
                                        MW.Show();
                                    }));
                                }
                                else
                                {
                                    MW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        MW.Show();
                                    }));
                                }
                                //if(debug_mode) stop_time();
                                grid_visible = true;
                                click_times = executions;
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.click)
                            {
                                LMBClick();
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.right_click)
                            {
                                RMBClick();
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.double_click)
                            {
                                DLMBClick();
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.triple_click)
                            {
                                TLMBClick();
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.hold)
                            {
                                LMBHold();
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.hold_right)
                            {
                                RMBHold();
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.release_buttons)
                            {
                                release_buttons();
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.move_top_edge)
                            {
                                int x = (int)(screen_width / 2);
                                int y = 0;

                                SetCursorPos(x, y);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.move_bottom_edge)
                            {
                                int x = (int)(screen_width / 2);
                                int y = screen_height - 1;

                                SetCursorPos(x, y);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.move_left_edge)
                            {
                                int y = (int)(screen_height / 2);
                                int x = 0;

                                SetCursorPos(x, y);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.move_right_edge)
                            {
                                int y = (int)(screen_height / 2);
                                int x = screen_width - 1;

                                SetCursorPos(x, y);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.move_screen_center)
                            {
                                int y = (int)(screen_height / 2);
                                int x = (int)(screen_width / 2);

                                SetCursorPos(x, y);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.scroll_up)
                            {
                                sim.Mouse.VerticalScroll(4);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.scroll_down)
                            {
                                sim.Mouse.VerticalScroll(-4);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.scroll_left)
                            {
                                sim.Mouse.HorizontalScroll(-30);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.scroll_right)
                            {
                                sim.Mouse.HorizontalScroll(30);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.close_that)
                            {
                                //left alt in WindowsInput library is bugged (keyup doesn't work)
                                //keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
                                //key_press(VirtualKeyCode.F4);
                                //keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                                //Thread.Sleep(1);

                                keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
                                key_down(VirtualKeyCode.F4);
                                Thread.Sleep(75);
                                key_up(VirtualKeyCode.F4);
                                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);

                                //Admin req (wasn't working before, but it's fixed now):
                                //IntPtr handle = GetForegroundWindow();

                                //Process[] arr = Process.GetProcesses();
                                //Process process = null;

                                //try
                                //{
                                //    foreach (Process p in arr)
                                //    {
                                //        IntPtr h = p.MainWindowHandle;

                                //        if (h == handle)
                                //        {
                                //            process = p;
                                //            process.CloseMainWindow();
                                //            break;
                                //        }
                                //    }
                                //}
                                //catch (Exception ex)
                                //{
                                //    try
                                //    {
                                //        if (process != null)
                                //            process.Kill(); //kill may not work if access is denied
                                //    }
                                //    catch (Exception ex2) { }
                                //}
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.minimize_that)
                            {
                                IntPtr handle = GetForegroundWindow();
                                ShowWindow(handle, SW_MINIMIZE);

                                /*
                                key_down(System.Windows.Forms.Keys.LWIN);
                                key_press(System.Windows.Forms.Keys.DOWN);
                                key_press(System.Windows.Forms.Keys.DOWN);
                                key_up(System.Windows.Forms.Keys.LWIN);
                                Thread.Sleep(1);
                                */
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.maximize_that)
                            {
                                IntPtr handle = GetForegroundWindow();
                                ShowWindow(handle, SW_MAXIMIZE);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.restore_that)
                            {
                                IntPtr handle = GetForegroundWindow();
                                ShowWindow(handle, SW_RESTORE);
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.get_position)
                            {
                                Middle_Man.last_get_position_point = System.Windows.Forms.Cursor.Position;

                                this.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    foreach (System.Windows.Window window in Application.Current.Windows)
                                    {
                                        if (window.GetType() == typeof(WindowAddEditActionMouse))
                                        {
                                            WindowAddEditActionMouse w = (WindowAddEditActionMouse)window;

                                            w.TBx.Text = Middle_Man.last_get_position_point.X.ToString();
                                            w.TBy.Text = Middle_Man.last_get_position_point.Y.ToString();
                                        }
                                        else if (window.GetType() == typeof(WindowAddEditActionMoveMouse))
                                        {
                                            WindowAddEditActionMoveMouse w = (WindowAddEditActionMoveMouse)window;

                                            w.TBx.Text = Middle_Man.last_get_position_point.X.ToString();
                                            w.TBy.Text = Middle_Man.last_get_position_point.Y.ToString();
                                        }
                                    }
                                }));
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.start_recording)
                            {
                                this.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    foreach (System.Windows.Window window in Application.Current.Windows)
                                    {
                                        if (window.GetType() == typeof(WindowRecordActions))
                                        {
                                            WindowRecordActions w = (WindowRecordActions)window;

                                            if (w.Bstart.Content.ToString() == "Start")
                                            {
                                                w.Bstart_Click(null, null);

                                                if (read_recognized_speech)
                                                {
                                                    if (ss.Volume != ss_volume)
                                                        ss.Volume = ss_volume;
                                                }
                                            }
                                        }
                                    }
                                }));
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.stop_recording)
                            {
                                this.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    foreach (System.Windows.Window window in Application.Current.Windows)
                                    {
                                        if (window.GetType() == typeof(WindowRecordActions))
                                        {
                                            WindowRecordActions w = (WindowRecordActions)window;

                                            if (w.Bstart.Content.ToString() != "Start")
                                            {
                                                w.Bstart_Click(null, null);

                                                if (read_recognized_speech)
                                                {
                                                    if (ss.Volume != ss_volume)
                                                        ss.Volume = ss_volume;
                                                }
                                            }
                                        }
                                    }
                                }));
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.switch_to_app)
                            {
                                Process[] arr = Process.GetProcesses();
                                string name = r.Replace(list_bic_general_and_mouse[i].name, "").Trim();
                                TimeSpan ts = new TimeSpan(0, 0, 0);
                                IntPtr handle = IntPtr.Zero;
                                string title;

                                foreach (Process p in arr)
                                {
                                    //if (p.MainWindowTitle.Contains(name) && p.TotalProcessorTime > ts)
                                    //title = p.MainWindowTitle.ToLower();
                                    title = p.MainWindowTitle;

                                    if (title.Contains(name))
                                    {
                                        //ts = p.TotalProcessorTime;
                                        handle = p.MainWindowHandle;
                                    }
                                }

                                if (handle != IntPtr.Zero)
                                {
                                    //ShowWindow(handle, SW_SHOWNORMAL);
                                    SetForegroundWindow(handle);
                                }

                                if (handle == IntPtr.Zero)
                                {
                                    foreach (Process p in arr)
                                    {
                                        if (p.ProcessName.Contains(name))
                                        {
                                            //length = p.ProcessName.Length;
                                            handle = p.MainWindowHandle;
                                        }
                                    }
                                }
                                if (handle != IntPtr.Zero)
                                {
                                    //ShowWindow(handle, SW_SHOWNORMAL);
                                    SetForegroundWindow(handle);
                                }
                            }
                            else if (list_bic_general_and_mouse[i].type == bic_type.open_app)
                            {
                                string name = r.Replace(list_bic_general_and_mouse[i].name, "").Trim();
                                bool app_found = false;

                                for (int k = 0; k < installed_apps.Count && app_found == false; k++)
                                {
                                    for (int m = 0; m < installed_apps[k].names.Count && app_found == false; m++)
                                    {
                                        if (name == installed_apps[k].names[m])
                                        {
                                            app_found = true;

                                            //was not working for some apps like Google Chrome until i changed
                                            //target platform from any to x64 in project settings
                                            Process.Start(installed_apps[k].path);

                                            //// IWshRuntimeLibrary is in the COM library "Windows Script Host Object Model"
                                            //IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();

                                            //try
                                            //{
                                            //    IWshRuntimeLibrary.IWshShortcut shortcut 
                                            //        = (IWshRuntimeLibrary.IWshShortcut)shell.{z
                                            //        CreateShortcut(installed_apps[k].path);
                                            //    Process.Start(System.IO.File.ReadAllText(shortcut.TargetPath));
                                            //}
                                            //catch (COMException)
                                            //{
                                            //    // A COMException is thrown if the file is not a valid shortcut (.lnk) file 
                                            //    MessageBox.Show("the file is not a valid shortcut (.lnk) file");
                                            //}
                                        }
                                    }
                                }
                            }
                            /*
                            else if (list_commands[i].type == bic_type.)
                            {

                            }
                            */
                        }
                    }

                    int range = 0;
                    int multiply = 1; //move by pixels nr is necessary for graphics work

                    if (found == false && n == 2 && int.TryParse(a[n - 1], out range))
                    {
                        if (a[0] == s_mouse_moves[0])
                        {
                            found = true;

                            real_move_mouse_by(0, range * -1 * multiply);
                        }
                        else if (a[0] == s_mouse_moves[1])
                        {
                            found = true;

                            real_move_mouse_by(0, range * multiply);
                        }
                        else if (a[0] == s_mouse_moves[2])
                        {
                            found = true;

                            real_move_mouse_by(range * -1 * multiply, 0);
                        }
                        else if (a[0] == s_mouse_moves[3])
                        {
                            found = true;

                            real_move_mouse_by(range * multiply, 0);
                        }
                    }

                    for (int i = 0; i < list_bic_keys_pressing.Count() && found == false; i++)
                    {
                        if (r == list_bic_keys_pressing[i].name && list_bic_keys_pressing[i].keys == null)
                        {
                            found = true;
                            if (executions == 1)
                            {
                                if (list_bic_keys_pressing[i].vkc == VirtualKeyCode.LMENU)
                                {
                                    keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
                                    Thread.Sleep(50);
                                    keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                                }
                                else
                                {
                                    key_press(list_bic_keys_pressing[i].vkc);
                                }
                            }
                            else
                            {
                                if (list_bic_keys_pressing[i].vkc == VirtualKeyCode.LMENU)
                                {
                                    keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
                                    Thread.Sleep(1);
                                    keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                                }
                                else
                                {
                                    key_press(list_bic_keys_pressing[i].vkc, 1);
                                }
                            }
                        }
                    }
                    //MessageBox.Show(found.ToString());
                    if (found == false)
                    {
                        int found_nr = 0;
                        List<VirtualKeyCode> codes = new List<VirtualKeyCode>();
                        
                        a = r.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                        int len = a.Length;

                        if (a.Contains("times"))
                            len -= 2;
                        else if (a.Contains("twice"))
                            len--;

                        if(len == 4)
                        {
                            a = new string[] { a[0] + " " + a[1], a[2] + " " + a[3] };
                        }
                        else if(len == 3)
                        {
                            bool f = false;

                            for (int i = 0; i < list_bic_keys_pressing.Count(); i++)
                            {
                                if(list_bic_keys_pressing[i].keys == a[0] + " " + a[1])
                                {
                                    f = true;
                                    a = new string[] { a[0] + " " + a[1], a[2] };
                                    break;
                                }
                            }
                            if(f == false)
                            {
                                for (int i = 0; i < list_bic_keys_pressing.Count(); i++)
                                {
                                    if (list_bic_keys_pressing[i].keys == a[1] + " " + a[2])
                                    {
                                        f = true;
                                        a = new string[] { a[0], a[1] + " " + a[2] };
                                        break;
                                    }
                                }
                            }
                        }

                        foreach (string key in a)
                        {   
                            for (int i = 0; i < list_bic_keys_pressing.Count() && found == false; i++)
                            {
                                if (key == list_bic_keys_pressing[i].name)
                                {
                                    found_nr++;
                                    codes.Add(list_bic_keys_pressing[i].vkc);
                                    break;
                                }
                            }
                        }
                        if (found_nr == 2)
                        {
                            found = true;

                            //MessageBox.Show(codes[0] + " : " + codes[1]);

                            key_down(codes[0]);

                            if (executions == 1)
                                key_press(codes[1]);
                            else
                                key_press(codes[1], 1);

                            key_up(codes[0]);
                        }
                    }
                    //MessageBox.Show(found.ToString());
                    for (int i = 0; i < list_bic_char_inserting.Count() && found == false; i++)
                    {
                        if (r == list_bic_char_inserting[i].name)
                        {   
                            found = true;
                            sim.Keyboard.TextEntry(list_bic_char_inserting[i].description);
                        }
                    }

                    if (control)
                        key_up(VirtualKeyCode.CONTROL);
                    if (shift)
                        key_up(VirtualKeyCode.SHIFT);
                    if (alt)
                        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                    if (windows)
                        key_up(VirtualKeyCode.LWIN);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW010", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            recognition_suspended = false;
        }

        private void save_grids()
        {
            FileStream fs = null;
            StreamWriter sw = null;
            
            try
            {
                string type_name = get_grid_folder_name_by_type(grid_type);

                string folder_path = Path.Combine(new string[] { Middle_Man.saving_folder_path, 
                    grids_foldername, type_name, desired_figures_nr.ToString()});

                if (Directory.Exists(folder_path) == false)
                {
                    Directory.CreateDirectory(folder_path);
                }

                //"Default Process_grid ind=0" is used to copy default grid elements
                //to new grids (for different apps)
                //that's why we don't save it

                for (int i = 1; i < grids.Count; i++)
                {
                    fs = new FileStream(System.IO.Path.Combine(new string[] {
                        folder_path, grids[i].process_name + ".txt" }), 
                        FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs);

                    sw.WriteLine(grids[i].process_name);
                    sw.WriteLine(grids[i].count);

                    foreach (Grid_element element in grids[i].elements)
                    {
                        sw.WriteLine(element.symbol);
                        sw.WriteLine(element.word);
                        sw.WriteLine(element.count);
                    }

                    sw.Close();
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW011", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex2) { }
            }
        }

        private void load_grids()
        {
            FileStream fs = null;
            StreamReader sr = null;
            
            try
            {
                string type_name = get_grid_folder_name_by_type(grid_type);

                string folder_path = Path.Combine(new string[] { Middle_Man.saving_folder_path,
                    grids_foldername, type_name, desired_figures_nr.ToString()});

                if (Directory.Exists(folder_path))
                {
                    string[] files = Directory.GetFiles(folder_path, "*.txt");
                    bool cancel = false;

                    //"Default Process_grid ind=0" is used to copy default grid elements
                    //to new grids (for different apps)
                    //that's why we don't remove first grid when we reset smart mousegrid data
                    //before loading new data

                    if (grids.Count > 1)
                    {
                        grids.RemoveRange(1, grids.Count - 1);
                    }

                    foreach (string file in files)
                    {
                        cancel = false;

                        fs = new FileStream(file, FileMode.Open, FileAccess.Read);
                        sr = new StreamReader(fs);

                        if (sr.EndOfStream)
                        {
                            sr.Close();
                            fs.Close();
                            continue;
                        }
                        Process_grid g = new Process_grid(sr.ReadLine());

                        if (sr.EndOfStream)
                        {
                            sr.Close();
                            fs.Close();
                            continue;
                        }
                        int count;
                        if (int.TryParse(sr.ReadLine(), out count) == false)
                        {
                            sr.Close();
                            fs.Close();
                            continue;
                        }
                        if (count != grids[grid_ind].count)
                        {
                            sr.Close();
                            fs.Close();
                            continue;
                        }
                        g.count = count;

                        g.elements = new List<Grid_element>();
                        string s, w;
                        uint count2;

                        while (sr.EndOfStream == false)
                        {
                            s = sr.ReadLine();

                            if (sr.EndOfStream)
                            {
                                cancel = true;
                                break;
                            }
                            w = sr.ReadLine();

                            if (sr.EndOfStream)
                            {
                                cancel = true;
                                break;
                            }
                            if (uint.TryParse(sr.ReadLine(), out count2) == false)
                            {
                                cancel = true;
                                break;
                            }

                            Grid_element element = new Grid_element(w, s);
                            element.count = count2;
                            g.elements.Add(element);
                        }

                        if (cancel == false)
                            grids.Add(g);

                        sr.Close();
                        fs.Close();
                    }
                }
                else //Support for Work by Speech v. 1.5 and earlier smart mousegrids data saving method
                {
                    folder_path = Path.Combine(new string[] { Middle_Man.saving_folder_path,
                        grids_foldername});

                    if (Directory.Exists(folder_path))
                    {
                        string[] files = Directory.GetFiles(folder_path, "*.txt");

                        folder_path = Path.Combine(new string[] { Middle_Man.saving_folder_path,
                        grids_foldername, type_name, desired_figures_nr.ToString()});

                        foreach (string file in files)
                        {
                            string[] arr = file.Split(new string[] { "\\" },
                                StringSplitOptions.RemoveEmptyEntries);
                            string filename = arr[arr.Length - 1];

                            if (Directory.Exists(folder_path) == false)
                            {
                                Directory.CreateDirectory(folder_path);
                            }

                            File.Move(file, Path.Combine(folder_path, filename));
                        }

                        if(files.Length > 0)
                            load_grids();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW012", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    sr.Close();
                    fs.Close();
                }
                catch (Exception ex2) { }
            }
        }

        void monitor()
        {
            while (true)
            {
                try
                {
                    if (SW.change_mode)
                    {
                        if (SW.mode == 0)
                        {
                            load_turned_off();
                        }
                        else if (SW.mode == 1)
                        {
                            current_mode = mode.command;
                            load_turned_on();
                        }
                        else
                        {
                            current_mode = mode.dictation;
                            load_turned_on();
                        }

                        SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                        {
                            SW.change_mode = false;
                        }));

                        Thread.Sleep(500);
                    }                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error MW013", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                Thread.Sleep(100);
            }
        }
    }
}