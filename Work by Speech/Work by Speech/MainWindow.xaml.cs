//highest error nr: MW013
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Windows.Threading;
using WindowsInput;
using WindowsInput.Native;
using System.Net;
using System.Globalization;
using System.Windows.Input;
using System.Windows.Data;

//Do NOT use asynchronous speech recognition, because after a couple of hours of using
//this app, it causes a bug when you have to say every command twice for it to be recognized

//If there are too many grammar builders these commands don't work: up 100, left 50, etc.
//because the dictionary is overloaded (so number of commands for each dictionary must
//be carefully controlled - it mostly concerns built-in commands).
//This app works well even after adding 1000 custom commands with max executions set to 50

//Long unloading grammars (sometimes 1 second, sometimes even 20 seconds) was caused by
//using a microphone connected via minijack while using an old sound card which
//doesn't officially support Windows 10. Using USB microphone solved the problem.

//ThreadPriority.Highest isn't really needed anywhere
//recognizer.RequestRecognizerUpdate(); isn't really needed anywhere
//initial_silence_timeout - defines the longest speech time that
//can be recognized (default 30 seconds)

//Fix cut window: x = -15px, y = -23px

namespace Speech
{
    public partial class MainWindow : Window
    {
        const string prog_version = "2.2-beta";
              string latest_version = "";
        const string copyright_text = "Copyright © 2023 - 2024 Mikołaj Magowski. All rights reserved.";
        const string filename_settings = "settings.xml";
        const string filename_coords = "coords.txt"; //speech recognition window last location
        const string grids_foldername = "grids";
        const int grid_symbols_limit = 50;//50-58 recommended
        const int max_font_size = 400;
        const bool resized_grid = true; //when true, resizes mousegrid so screen is fully covered
        const bool movable_grid = true; //when true, mousegrid can be moved by speech

        const bool dont_unload_grammars = false; //set this to true if grammar unloading would ever become
                                                 //bugged (need to update some code to use this)
        
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

        SpeechRecognitionEngine recognizer;

        Grammar grammar_off_mode;
        Grammar grammar_dictation_commands;
        Grammar grammar_dictation;
        Grammar grammar_mousegrid;
        Grammar grammar_builtin_commands;
        Grammar grammar_custom_commands_any; //any program
        Grammar grammar_custom_commands_foreground; //foreground program
        Grammar grammar_apps_switching;
        Grammar grammar_apps_opening;

        const string grammar_off_mode_name = "off mode";
        const string grammar_dictation_commands_name = "dictation commands";
        const string grammar_dictation_name = "dictation";
        const string grammar_mousegrid_name = "mousegrid";
        const string grammar_builtin_commands_name = "built-in commands";
        const string grammar_custom_commands_any_name = "custom commands any";
        const string grammar_custom_commands_foreground_name = "custom commands foreground";
        const string grammar_apps_switching_name = "apps switching";
        const string grammar_apps_opening_name = "apps opening";

        List<string> dictation_commands;

        //---------------Main voice commands START---------------
        const string turn_on = "start speech recognition";
        const string turn_off = "recognition off";
        const string switch_to_command_mode = "command";
        const string switch_to_dictation_mode = "dictation";
        const string show_speech_recognition = "show speech recognition"; //window
        const string hide_speech_recognition = "hide speech recognition";
        const string start_better_dictation_listening = "start listening"; //used in better dictation
        const string toggle_better_dictation_str = "toggle"; //used in better dictation
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
        Thread THRrecognition;
        Thread THRholder; //for holding keys
        //bool thread_abort1 = false; //redeclaring and starting stopped thread doesn't work (weird)
        bool thread_suspend1 = false;
        bool thread_suspended1 = false;
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
        const bool default_better_dictation = false; //if set to true, use windows dictation tool
                                                    //(Windows key + H)
        
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
        int confidence_dictation;
        bool better_dictation; //if set to true, use windows dictation tool (Windows key + H)

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

                //Access All Users Start Menu
                StringBuilder path = new StringBuilder(260);
                SHGetSpecialFolderPath(IntPtr.Zero, path, CSIDL_COMMON_STARTMENU, false);
                start_menu_path = path.ToString();

                try
                {
                    Directory.GetFiles(start_menu_path, "*.*", SearchOption.AllDirectories);
                }
                catch (Exception ex)
                {
                    //SHGetSpecialFolderPath may cause a bug if some folders in "C:\\ProgramData\\Microsoft\\Windows\\Start Menu"
                    //path that it returns have denied access
                    start_menu_path = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs";

                    try
                    {
                        Directory.GetFiles(start_menu_path, "*.*", SearchOption.AllDirectories);
                    }
                    catch (Exception ex2)
                    {
                        start_menu_path = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programmes";

                        try
                        {
                            Directory.GetFiles(start_menu_path, "*.*", SearchOption.AllDirectories);
                        }
                        catch (Exception ex3)
                        {
                            //less shortcuts unfortunately
                            start_menu_path = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);
                        }
                    }
                }
                
                //not enough shortcuts:
                //start_menu_path = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);

                is_program_already_running();

                InitializeComponent();

                this.Title = Middle_Man.prog_name;
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
                Benable_bic_always.IsEnabled = false;
                Bdisable_bic_always.IsEnabled = false;
                Benable_bic_better.IsEnabled = false;
                Bdisable_bic_better.IsEnabled = false;

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
                MIenable_bic_always.IsEnabled = false;
                MIdisable_bic_always.IsEnabled = false;
                MIenable_bic_better.IsEnabled = false;
                MIdisable_bic_better.IsEnabled = false;

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
                        "will be restored to default and saved.", "Warning",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
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

                string culture = CultureInfo.CurrentUICulture.Name;

                try
                {
                    // Create an in-process speech recognizer for the en-US locale.  
                    //recognizer = new SpeechRecognitionEngine(new CultureInfo("en-US"));
                    //recognizer = new SpeechRecognitionEngine(new CultureInfo("en"));
                    recognizer = new SpeechRecognitionEngine(new CultureInfo(culture));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("This program requires English (United States, " +
                        "United Kingdom, Canada, India or Australia) set as default Windows language." +
                        "\n\nIn order to change your Windows display language go to:" +
                        "\nStart -> Settings -> Time & Language -> Language.",
                        "Error MW002", MessageBoxButton.OK, MessageBoxImage.Error);
                    Process.GetCurrentProcess().Kill();
                }

                // Add a handler for the speech recognized event.  
                recognizer.SpeechRecognized +=
                      new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);

                // Configure input to the speech recognizer.
                try
                {
                    recognizer.SetInputToDefaultAudioDevice();
                }
                catch (InvalidOperationException ex5)
                {
                    MessageBox.Show("Please connect a microphone to your computer before running this " +
                        "application.",
                            "Error MW003", MessageBoxButton.OK, MessageBoxImage.Error);
                    Process.GetCurrentProcess().Kill();
                }
                // Configure recognition parameters.  
                //Default (year 2022):
                //recognizer.InitialSilenceTimeout = TimeSpan.FromSeconds(30.0);
                //recognizer.BabbleTimeout = TimeSpan.FromSeconds(0);
                //recognizer.EndSilenceTimeout = TimeSpan.FromSeconds(0.15);
                //recognizer.EndSilenceTimeoutAmbiguous = TimeSpan.FromSeconds(0.5);

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

                create_grammars();

                //start_time2();
                load_grammars(true, false);
                //stop_time2();

                load_turned_off();

                if (smart_grid)
                    load_grids();

                // Start asynchronous, continuous speech recognition.  
                //recognizer.RecognizeAsync(RecognizeMode.Multiple); //sometimes picks too much speech
                //especially after running for a long time

                //Best solution for speech recognition:
                THRrecognition = new Thread(new ThreadStart(speech_recognition));
                THRrecognition.Start();

                if (CHBstart_with_hidden.IsChecked == false)
                {
                    SW.Show();
                }

                THRmonitor = new Thread(new ThreadStart(monitor));
                THRmonitor.Start();

                THRholder = new Thread(new ThreadStart(hold_keys_and_buttons));
                THRholder.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW004", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            Mouse.OverrideCursor = null;
        }

        void speech_recognition()
        {
            while(true)
            {
                SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                {
                    SW.Bmode.Visibility = Visibility.Visible;
                }));

                if (dont_unload_grammars)
                {
                    try
                    {
                        recognizer.Recognize();
                    }
                    catch (Exception ex)
                    {
                        //fu
                    }
                }
                else
                    recognizer.Recognize();

                //recognizer.RecognizeAsync(RecognizeMode.Single);
                //while (speech_recognized == false)
                //{
                //    Thread.Sleep(10);
                //}
                //speech_recognized = false;

                SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                {
                    SW.Bmode.Visibility = Visibility.Hidden;
                }));

                Thread.Sleep(10);

                //inside_speech_recognized_event isn't really necessary, because Recognize() without
                //async means that this thread will execute speech recognized event
                while (recognition_suspended || inside_speech_recognized_event)
                    //|| ss.State == SynthesizerState.Speaking) //slows down too much
                {
                    Thread.Sleep(10);
                }

                Thread.Sleep(10); //at least 3ms, because of how fast sound travels
            }
        }

        void create_grammars()
        {
            if (are_all_bic_off_disabled() == false)
                grammar_off_mode = create_off_mode_grammar();
            else
                grammar_off_mode = null;

            grammar_mousegrid = create_grid_grammar(true);

            if (are_all_bic_dictation_disabled() == false)
                grammar_dictation_commands = create_dictation_commands_grammar();
            else
                grammar_dictation_commands = null;

            // Create and load a dictation grammar.
            grammar_dictation = new DictationGrammar();
            //grammar_dictation = new DictationGrammar("grammar:dictation");
            //grammar_dictation = new DictationGrammar("grammar:dictation#spelling");
            grammar_dictation.Name = grammar_dictation_name;

            if (are_all_bic_general_and_mouse_disabled() == false)
                grammar_builtin_commands = create_builtin_commands_grammar();
            else
                grammar_builtin_commands = null;

            grammar_custom_commands_any = create_custom_commands_grammar(Middle_Man.any_program_name);
        }

        void load_grammars(bool first_execute = false, bool reset_smart_grid = true)
        {
            if (grammar_off_mode != null)
                recognizer.LoadGrammar(grammar_off_mode);

            if (reset_smart_grid)
            {
                grammar_mousegrid = create_grid_grammar(true);
            }

            if (grammar_mousegrid != null)
                recognizer.LoadGrammar(grammar_mousegrid);

            if (grammar_dictation_commands != null)
                recognizer.LoadGrammar(grammar_dictation_commands);

            if (grammar_dictation != null)
                recognizer.LoadGrammar(grammar_dictation);

            if (grammar_builtin_commands != null)
                recognizer.LoadGrammar(grammar_builtin_commands);

            if(grammar_custom_commands_any != null)
                recognizer.LoadGrammar(grammar_custom_commands_any);

            //Debug only:
            //display_grammars_status();

            if (first_execute == true)
            {
                single_app_grammar_update(true);
            }
            else //can't happen anymore (load_grammars is currently executed only once)
            {
                Middle_Man.force_updating_both_cc_grammars = true;

                if (grammar_apps_switching != null & apps_switching)
                    recognizer.LoadGrammar(grammar_apps_switching);
                if (grammar_apps_opening != null && apps_opening)
                    recognizer.LoadGrammar(grammar_apps_opening);
            }

            if (first_execute == true)
            {
                THRswitch_to = new Thread(new ThreadStart(keep_grammars_updated));
                THRswitch_to.Start();
            }
        }

        void restart_recognizer(bool reset_smart_grid)
        {
            recognition_suspended = true;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Hidden;
            }));

            //THRswitch_to.Abort(); //abort only causes problems
            thread_suspend1 = true;

            //wait for suspend:
            while (thread_suspended1 == false)
            {
                Thread.Sleep(10);
            }

            List<bool> grammars_status = new List<bool>();

            for(int i=0; i<recognizer.Grammars.Count; i++)
            {
                grammars_status.Add(recognizer.Grammars[i].Enabled);
            }

            start_time2();
            recognizer.Dispose();
            stop_time2();

            string culture = CultureInfo.CurrentUICulture.Name;
            
            try
            {
                recognizer = new SpeechRecognitionEngine(new CultureInfo(culture));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\nPlease change your Windows display language" +
                    " to English (United States, United Kingdom, Canada, India or Australia).",
                    "Error MW005", MessageBoxButton.OK, MessageBoxImage.Error);
                Process.GetCurrentProcess().Kill();
            }

            recognizer.SpeechRecognized +=
                  new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);

            try
            {
                recognizer.SetInputToDefaultAudioDevice();
            }
            catch (InvalidOperationException ex5)
            {
                MessageBox.Show("Please connect a microphone to your computer before running this " +
                    "application.",
                        "Error MW006", MessageBoxButton.OK, MessageBoxImage.Error);
                Process.GetCurrentProcess().Kill();
            }

            load_grammars(false, reset_smart_grid);

            for (int i = 0; i < grammars_status.Count; i++)
            {
                recognizer.Grammars[i].Enabled = grammars_status[i];
            }

            thread_suspend1 = false;

            recognition_suspended = false;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Visible;
            }));
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

        string prev_installed_apps_str = "";

        bool get_installed_apps()
        {
            string[] allfiles = Directory.GetFiles(start_menu_path, "*.*", SearchOption.AllDirectories);
            
            string[] a;
            string str;
            List<string> names;
            string installed_app_path, s;
            int name_length;
            Installed_App installed_app;
            string installed_apps_str = "";

            List<Installed_App> installed_apps2 = new List<Installed_App>();

            foreach (var file in allfiles)
            {
                if (thread_suspend1)
                    return false;

                FileInfo info = new FileInfo(file);

                installed_apps_str += info.Name;
            }

            if (installed_apps_str == prev_installed_apps_str)
            {
                return false;
            }

            foreach (var file in allfiles)
            {
                if (thread_suspend1)
                    return false;

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
                if (thread_suspend1)
                    return false;
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

            if (thread_suspend1)
                return false;
            else
            {
                installed_apps = installed_apps2;
                prev_installed_apps_str = installed_apps_str;
                return true;
            }
        }

        void load_turned_off()
        {
            int i = 0;
            
            recognition_suspended = true;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Hidden;
            }));

            //THRswitch_to.Abort(); //abort only causes problems
            thread_suspend1 = true;
            
            //wait for suspend:
            while (thread_suspended1 == false)
            {
                Thread.Sleep(10);
                i++;
            }

            toggle_grammar(true, grammar_type.grammar_off_mode);
            toggle_grammar(false, grammar_type.grammar_mousegrid);
            toggle_grammar(false, grammar_type.grammar_dictation_commands);
            toggle_grammar(false, grammar_type.grammar_dictation);
            toggle_grammar(false, grammar_type.grammar_builtin_commands);
            toggle_grammar(false, grammar_type.grammar_custom_commands_any);
            toggle_grammar(false, grammar_type.grammar_custom_commands_foreground);
            toggle_grammar(false, grammar_type.grammar_apps_switching);
            toggle_grammar(false, grammar_type.grammar_apps_opening);

            //Debug only:
            //display_grammars_status();

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

            thread_suspend1 = false;

            if (current_mode != mode.off)
                last_mode = current_mode;

            current_mode = mode.off;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Visible;
            }));
        }

        void load_turned_on(bool allow_dictation_toggle = true)
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
            thread_suspend1 = true;
            
            //wait for suspend:
            while (thread_suspended1 == false)
            {
                Thread.Sleep(10);
                i++;
            }

            toggle_grammar(false, grammar_type.grammar_off_mode);
            toggle_grammar(false, grammar_type.grammar_mousegrid);
            
            if (current_mode == mode.command)
            {
                toggle_grammar(false, grammar_type.grammar_dictation_commands);
                toggle_grammar(false, grammar_type.grammar_dictation);

                toggle_grammar(true, grammar_type.grammar_builtin_commands);
                toggle_grammar(true, grammar_type.grammar_custom_commands_any);
                toggle_grammar(true, grammar_type.grammar_custom_commands_foreground);
                
                if (is_bic_in_general_and_mouse_enabled(bic_type.switch_to_app))
                    toggle_grammar(true, grammar_type.grammar_apps_switching);
                else
                    toggle_grammar(false, grammar_type.grammar_apps_switching);
                
                if (is_bic_in_general_and_mouse_enabled(bic_type.open_app))
                    toggle_grammar(true, grammar_type.grammar_apps_opening);
                else
                    toggle_grammar(false, grammar_type.grammar_apps_opening);

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
            }
            else if(current_mode == mode.dictation)
            {
                toggle_grammar(true, grammar_type.grammar_dictation_commands);
                toggle_grammar(true, grammar_type.grammar_dictation);
                toggle_grammar(false, grammar_type.grammar_builtin_commands);
                toggle_grammar(false, grammar_type.grammar_custom_commands_any);
                toggle_grammar(false, grammar_type.grammar_custom_commands_foreground);
                toggle_grammar(false, grammar_type.grammar_apps_switching);
                toggle_grammar(false, grammar_type.grammar_apps_opening);

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

                if (better_dictation && allow_dictation_toggle)
                    toggle_better_dictation();
            }

            thread_suspend1 = false;

            recognition_suspended = false;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Visible;
            }));
        }

        private void enable_grid_gr()
        {
            int i = 0;
            
            recognition_suspended = true;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Hidden;
            }));

            //THRswitch_to.Abort(); //abort only causes problems
            thread_suspend1 = true;
            
            //wait for suspend:
            while (thread_suspended1 == false)
            {
                Thread.Sleep(10);
                i++;
            }

            toggle_grammar(false, grammar_type.grammar_off_mode);
            toggle_grammar(true, grammar_type.grammar_mousegrid);
            toggle_grammar(false, grammar_type.grammar_dictation_commands);
            toggle_grammar(false, grammar_type.grammar_dictation);
            toggle_grammar(false, grammar_type.grammar_builtin_commands);
            toggle_grammar(false, grammar_type.grammar_custom_commands_any);
            toggle_grammar(false, grammar_type.grammar_custom_commands_foreground);
            toggle_grammar(false, grammar_type.grammar_apps_switching);
            toggle_grammar(false, grammar_type.grammar_apps_opening);
                        
            thread_suspend1 = false;

            recognition_suspended = false;

            current_mode = mode.grid;

            SW.Bmode.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
            {
                SW.Bmode.Visibility = Visibility.Visible;
            }));
        }

        string s_last_ch = "";
        void keep_grammars_updated()
        {
            while (true)
            {
                try
                {
                    //is_command_mode_active() is not enough, because all grammars are temporarily
                    //enabled when load_grammars() is executed
                    if (current_mode == mode.command && THRrecognition != null)
                    {
                        single_app_grammar_update();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error MW008", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                for (int i = 0; i < 50 && thread_suspend1 == false; i++)
                {
                    Thread.Sleep(10);
                }

                thread_suspended1 = true;

                while (thread_suspend1)
                {
                    Thread.Sleep(50);
                }

                thread_suspended1 = false;
            }
        }

        string last_foreground_window_process = "";

        void single_app_grammar_update(bool first_execute = false)
        {
            bool changes_detected = false;
            string s = "", w, s_ch = "";

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

            Grammar new_grammar_custom_commands_foreground = create_custom_commands_grammar(process_name);

            if (Middle_Man.force_updating_both_cc_grammars)
            {
                if (dont_unload_grammars == false)
                {
                    if (is_grammar_loaded(grammar_custom_commands_foreground_name))
                        recognizer.UnloadGrammar(grammar_custom_commands_foreground);

                    if (is_grammar_loaded(grammar_custom_commands_any_name))
                        recognizer.UnloadGrammar(grammar_custom_commands_any);
                }

                grammar_custom_commands_foreground = new_grammar_custom_commands_foreground;
                grammar_custom_commands_any = create_custom_commands_grammar(Middle_Man.any_program_name);

                if (dont_unload_grammars == false)
                {
                    //recognizer.RequestRecognizerUpdate();
                    if (grammar_custom_commands_any != null)
                    {
                        recognizer.LoadGrammar(grammar_custom_commands_any);
                        toggle_grammar(true, grammar_type.grammar_custom_commands_any);
                    }
                        
                    if (grammar_custom_commands_foreground != null)
                    {
                        recognizer.LoadGrammar(grammar_custom_commands_foreground);
                        toggle_grammar(true, grammar_type.grammar_custom_commands_foreground);
                    }
                }

                last_foreground_window_process = process_name;

                Middle_Man.force_updating_both_cc_grammars = false;
            }
            else if (last_foreground_window_process != process_name)
            {
                if (dont_unload_grammars == false)
                {
                    if (is_grammar_loaded(grammar_custom_commands_foreground_name))
                        recognizer.UnloadGrammar(grammar_custom_commands_foreground);
                }

                grammar_custom_commands_foreground = new_grammar_custom_commands_foreground;

                if (new_grammar_custom_commands_foreground != null)
                {
                    //recognizer.RequestRecognizerUpdate();
                    recognizer.LoadGrammar(grammar_custom_commands_foreground);
                    toggle_grammar(true, grammar_type.grammar_custom_commands_foreground);
                }

                last_foreground_window_process = process_name;
            }
            
            if (apps_switching)
            {
                s = switch_to_app_str;

                Choices ch = new Choices();
                GrammarBuilder gb = new GrammarBuilder();
                gb.Append(s);

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
                                ch.Add(new string[] { a[i] });
                                s_ch += a[i];
                            }
                        }
                        for (int i = 0; i < a.Length - 1; i++)
                        {
                            a[i] = a[i].Trim();
                            a[i + 1] = a[i + 1].Trim();

                            if (a[i].Length > 0 || a[i + 1].Length > 0)
                            {
                                w = a[i] + " " + a[i + 1];
                                ch.Add(new string[] { w });
                                s_ch += w;
                            }
                        }
                        if (a.Length > 2)
                        {
                            ch.Add(new string[] { name });
                            s_ch += name;
                        }
                    }

                    if (thread_suspend1)
                        break;
                }

                if (thread_suspend1 == false)
                {
                    //usually 30ms
                    if (s_last_ch != s_ch)
                    {
                        changes_detected = true;

                        gb.Append(ch);

                        if (dont_unload_grammars == false)
                        {
                            //start_time2();
                            //recognizer.RequestRecognizerUpdate();
                            if (is_grammar_loaded(grammar_apps_switching_name))
                                recognizer.UnloadGrammar(grammar_apps_switching);
                            //stop_time2();
                        }

                        grammar_apps_switching = new Grammar(gb);
                        grammar_apps_switching.Name = grammar_apps_switching_name;

                        if (dont_unload_grammars == false && grammar_apps_switching != null)
                        {
                            //recognizer.RequestRecognizerUpdate();
                            recognizer.LoadGrammar(grammar_apps_switching);
                            toggle_grammar(true, grammar_type.grammar_apps_switching);
                        }

                        s_last_ch = s_ch;
                    }
                }
            }

            //dt1 = DateTime.Now;
            //dt2 = DateTime.Now;
            //ts = dt2 - dt1;
            //MessageBox.Show("processes: " + ts.TotalMilliseconds.ToString());

            //if get_installed_apps() == true means apps number has changed
            if (apps_opening && thread_suspend1 == false && get_installed_apps())
            {
                Choices ch = new Choices();
                GrammarBuilder gb = new GrammarBuilder();

                changes_detected = true;

                s = open_app_str;

                ch = new Choices();
                gb = new GrammarBuilder();
                gb.Append(s);

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
                                ch.Add(new string[] { name2 });
                            }
                        }
                    }
                }

                gb.Append(ch);

                if (dont_unload_grammars == false && first_execute == false)
                {
                    //recognizer.RequestRecognizerUpdate();
                    if (is_grammar_loaded(grammar_apps_opening_name))
                        recognizer.UnloadGrammar(grammar_apps_opening);
                }

                grammar_apps_opening = new Grammar(gb);
                grammar_apps_opening.Name = grammar_apps_opening_name;

                if (dont_unload_grammars == false && grammar_apps_opening != null)
                {
                    //recognizer.RequestRecognizerUpdate();
                    recognizer.LoadGrammar(grammar_apps_opening);
                    toggle_grammar(true, grammar_type.grammar_apps_opening);
                }
            }

            if (dont_unload_grammars && first_execute == false && changes_detected)
            {
                restart_recognizer(false);
            }
        }

        Grammar create_off_mode_grammar()
        {
            Grammar gr = new Grammar(new GrammarBuilder(turn_on));
            gr.Name = grammar_off_mode_name;

            return gr;
        }

        Grammar create_dictation_commands_grammar()
        {
            Choices ch = new Choices();
            GrammarBuilder gb;

            for (int i = 0; i < list_bic_dictation_always.Count(); i++)
            {
                if (list_bic_dictation_always[i].enabled)
                {
                    gb = new GrammarBuilder();
                    gb.Append(list_bic_dictation_always[i].name);
                    ch.Add(gb);
                }
            }

            if (better_dictation)
            {
                for (int i = 0; i < list_bic_dictation_better.Count(); i++)
                {
                    if (list_bic_dictation_better[i].enabled)
                    {
                        gb = new GrammarBuilder();
                        gb.Append(list_bic_dictation_better[i].name);
                        ch.Add(gb);
                    }
                }
            }

            Grammar gr = new Grammar(ch);
            gr.Name = grammar_dictation_commands_name;

            return gr;
        }

        Grammar create_grid_grammar(bool reset_smart_grid)
        {
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

            GrammarBuilder gb;
            Choices ch = new Choices();

            Choices cancels = new Choices(cancels_str);
            Choices directions = new Choices(directions_str);
            Choices drag_edges = new Choices(drag_edges_str);

            gb = new GrammarBuilder();
            gb.Append(cancels);
            ch.Add(gb);

            gb = new GrammarBuilder();
            gb.Append(directions);
            ch.Add(gb);

            gb = new GrammarBuilder();
            gb.Append(drag_edges);
            ch.Add(gb);

            for (int i = 0; i < count; i++)
            {
                gb = new GrammarBuilder();
                gb.Append(grid_alphabet[i].word);
                ch.Add(gb);
            }
            for (int i = 0; i < count; i++)
            {
                for (int j = 0; j < count; j++)
                {
                    gb = new GrammarBuilder();
                    gb.Append(grid_alphabet[i].word);
                    gb.Append(grid_alphabet[j].word);
                    ch.Add(gb);
                }
            }
            for (int i = 0; i < count; i++)
            {
                for (int j = 0; j < count; j++)
                {
                    if (grid_alphabet[i].word == grid_alphabet[j].word)
                    {
                        gb = new GrammarBuilder();
                        gb.Append(grid_alphabet[i].word + " twice");
                        ch.Add(gb);
                    }
                }
            }

            int r1, r2;
            for (int i = 0; i < count; i++)
            {
                for (int j = 0; j < count; j++)
                {
                    if (int.TryParse(grid_alphabet[i].symbol, out r1)
                        && int.TryParse(grid_alphabet[j].symbol, out r2))
                    {
                        gb = new GrammarBuilder();
                        gb.Append(grid_alphabet[i].symbol + grid_alphabet[j].symbol);
                        ch.Add(gb);
                    }
                }
            }

            Grammar grammar = new Grammar(ch);
            grammar.Name = grammar_mousegrid_name;
            return grammar;
        }

        Grammar create_builtin_commands_grammar()
        {
            GrammarBuilder gb;
            Choices ch = new Choices();

            Choices pixels = new Choices();

            Choices combo = new Choices(s_combo);
            Choices combo2 = new Choices(s_combo2);
            Choices multi = new Choices(new string[] { "twice" });

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
                    gb = new GrammarBuilder();
                    gb.Append(list_bic_general_and_mouse[i].name);
                    ch.Add(gb);

                    if (list_bic_general_and_mouse[i].key_combination == "Yes")
                    {
                        gb = new GrammarBuilder();
                        gb.Append(combo);
                        gb.Append(list_bic_general_and_mouse[i].name);
                        ch.Add(gb);
                        gb = new GrammarBuilder();
                        gb.Append(combo);
                        gb.Append(combo2);
                        gb.Append(list_bic_general_and_mouse[i].name);
                        ch.Add(gb);
                    }

                    if (list_bic_general_and_mouse[i].max_executions > 1)
                    {
                        multi = new Choices(new string[] { "twice" });

                        for (int j = 2; j <= list_bic_general_and_mouse[i].max_executions; j++)
                        {
                            multi.Add(new string[] { j + " times" });
                        }

                        gb = new GrammarBuilder();
                        gb.Append(list_bic_general_and_mouse[i].name);
                        gb.Append(multi);
                        ch.Add(gb);
                    }

                    if (list_bic_general_and_mouse[i].key_combination == "Yes" 
                        && list_bic_general_and_mouse[i].max_executions > 1)
                    {
                        gb = new GrammarBuilder();
                        gb.Append(combo);
                        gb.Append(list_bic_general_and_mouse[i].name);
                        gb.Append(multi);
                        ch.Add(gb);

                        gb = new GrammarBuilder();
                        gb.Append(combo);
                        gb.Append(combo2);
                        gb.Append(list_bic_general_and_mouse[i].name);
                        gb.Append(multi);
                        ch.Add(gb);
                    }
                }
            }

            //Choices ch2 = new Choices();
            //for (int i = 0; i < list_bic_keys_pressing.Count(); i++)
            //{
            //    if (list_bic_keys_pressing[i].key_combination == "Yes")
            //    {
            //        ch2.Add(list_bic_keys_pressing[i].name);
            //    }
            //}

            //Choices ch3 = new Choices();
            //for (int i = 0; i < list_bic_keys_pressing.Count(); i++)
            //{
            //    if (list_bic_keys_pressing[i].key_combination == "Yes"
            //      && list_bic_keys_pressing[i].max_executions > 1)
            //    {
            //        ch3.Add(list_bic_keys_pressing[i].name);
            //    }
            //}

            for (int i = 0; i < list_bic_keys_pressing.Count(); i++)
            {
                if (list_bic_keys_pressing[i].enabled)
                {
                    gb = new GrammarBuilder();
                    gb.Append(list_bic_keys_pressing[i].name);
                    ch.Add(gb);

                    if (list_bic_keys_pressing[i].key_combination == "Yes")
                    {
                        gb = new GrammarBuilder();
                        gb.Append(combo);
                        gb.Append(list_bic_keys_pressing[i].name);
                        ch.Add(gb);

                        gb = new GrammarBuilder();
                        gb.Append(combo);
                        gb.Append(combo);
                        gb.Append(list_bic_keys_pressing[i].name);
                        ch.Add(gb);
                    }

                    if (list_bic_keys_pressing[i].max_executions > 1)
                    {
                        multi = new Choices(new string[] { "twice" });

                        for (int j = 2; j <= list_bic_keys_pressing[i].max_executions; j++)
                        {
                            multi.Add(new string[] { j + " times" });
                        }

                        gb = new GrammarBuilder();
                        gb.Append(list_bic_keys_pressing[i].name);
                        gb.Append(multi);
                        ch.Add(gb);
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

            if (list_bic_keys_pressing[i].key_combination == "Yes" 
                        && list_bic_keys_pressing[i].max_executions > 1)
                    {
                        gb = new GrammarBuilder();
                        gb.Append(combo);
                        gb.Append(list_bic_keys_pressing[i].name);
                        gb.Append(multi);
                        ch.Add(gb);

                        gb = new GrammarBuilder();
                        gb.Append(combo);
                        gb.Append(combo);
                        gb.Append(list_bic_keys_pressing[i].name);
                        gb.Append(multi);
                        ch.Add(gb);
                    }
                }
            }
            
            for (int i = 0; i < list_bic_char_inserting.Count(); i++)
            {
                if (list_bic_char_inserting[i].enabled)
                {
                    gb = new GrammarBuilder();
                    gb.Append(list_bic_char_inserting[i].name);
                    ch.Add(gb);

                    if (list_bic_char_inserting[i].max_executions > 1)
                    {
                        multi = new Choices(new string[] { "twice" });

                        for (int j = 2; j <= list_bic_char_inserting[i].max_executions; j++)
                        {
                            multi.Add(new string[] { j + " times" });
                        }

                        gb = new GrammarBuilder();
                        gb.Append(list_bic_char_inserting[i].name);
                        gb.Append(multi);
                        ch.Add(gb);
                    }
                }
            }

            //careful - adding too many can overload the dictionary
            for (int i = 1; i <= 100; i++)
            {
                pixels.Add(new string[] { i.ToString() });
            }
            //for (int i = 100; i <= 8000; i += 50)
            //for (int i = 150; i <= 1000; i += 50)
            //{
            //    pixels.Add(new string[] { i.ToString() });
            //}

            List<string> enabled_mouse_moves = new List<string>();

            if (is_bic_in_general_and_mouse_enabled(bic_type.move_up))
                enabled_mouse_moves.Add(s_mouse_moves[0]);
            if (is_bic_in_general_and_mouse_enabled(bic_type.move_down))
                enabled_mouse_moves.Add(s_mouse_moves[1]);
            if (is_bic_in_general_and_mouse_enabled(bic_type.move_left))
                enabled_mouse_moves.Add(s_mouse_moves[2]);
            if (is_bic_in_general_and_mouse_enabled(bic_type.move_right))
                enabled_mouse_moves.Add(s_mouse_moves[3]);

            Choices mouse_moves = new Choices(enabled_mouse_moves.ToArray());

            if (enabled_mouse_moves.Count > 0)
            {
                gb = new GrammarBuilder();
                gb.Append(mouse_moves);
                gb.Append(pixels);
                ch.Add(gb);
            }

            Grammar grammar = new Grammar(ch);
            grammar.Name = grammar_builtin_commands_name;
            return grammar;
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

        Grammar create_custom_commands_grammar(string program)
        {
            GrammarBuilder gb;
            Choices ch = new Choices();

            Choices multi = new Choices(new string[] { "twice" });

            program = program.ToLower();

            int cc2_commands_found = 0;

            //adding same words to grammar multiple times may generate exception:
            //System.FormatException: ''' rule reference not defined in this grammar.'
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
                                gb = new GrammarBuilder();
                                gb.Append(cc.name);
                                ch.Add(gb);

                                CustomCommand u = new CustomCommand();
                                u.name = cc.name;
                                u.max_executions = cc.max_executions;

                                unique_commands.Add(u);
                            }

                            if (found_both == false && cc.max_executions > 1)
                            {
                                multi = new Choices(new string[] { "twice" });

                                for (int j = 2; j <= cc.max_executions; j++)
                                {
                                    multi.Add(new string[] { j + " times" });
                                }

                                gb = new GrammarBuilder();
                                gb.Append(cc.name);
                                gb.Append(multi);
                                ch.Add(gb);

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

            Grammar grammar = null;

            if (cc2_commands_found > 0)
            {
                grammar = new Grammar(ch);

                if (program == Middle_Man.any_program_name.ToLower())
                    grammar.Name = grammar_custom_commands_any_name;
                else
                    grammar.Name = grammar_custom_commands_foreground_name;
            }            

            return grammar;
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
                if ((r == drag_edges_str[0] || r == drag_edges_str[1] 
                    || r == drag_edges_str[2] || r == drag_edges_str[3]
                    || r == drag_edges_str[4]))
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

                    if (r == drag_edges_str[0])
                    {
                        x = (int)(screen_width / 2);
                        y = 0;
                    }
                    else if (r == drag_edges_str[1])
                    {
                        x = (int)(screen_width / 2);
                        y = screen_height - 1;
                    }
                    else if (r == drag_edges_str[2])
                    {
                        y = (int)(screen_height / 2);
                        x = 0;
                    }
                    else if (r == drag_edges_str[3])
                    {
                        y = (int)(screen_height / 2);
                        x = screen_width - 1;
                    }
                    else if (r == drag_edges_str[4])
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
                enable_grid_gr();
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

        string r, r_lowercase;
        bool speech_recognized = false;

        // Handle the SpeechRecognized event.  
        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            speech_recognized = true;
            inside_speech_recognized_event = true;

            try
            {
                if (recognition_suspended == false)
                {
                    if (ss.Volume != ss_volume)
                        ss.Volume = ss_volume;

                    r = e.Result.Text;
                    r_lowercase = r.ToLower();
                    int c = (int)Math.Floor(e.Result.Confidence * 100);

                    bool dictation_command = false;

                    foreach (BuiltInCommand bic in list_bic_dictation_always)
                    {
                        if (r == bic.name && bic.enabled)
                        {
                            dictation_command = true;
                            break;
                        }
                    }

                    if (better_dictation)
                    {
                        foreach (BuiltInCommand bic in list_bic_dictation_better)
                        {
                            if (r == bic.name && bic.enabled)
                            {
                                dictation_command = true;
                                break;
                            }
                        }
                    }

                    SW.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                    {
                        if (current_mode != mode.dictation || better_dictation == false
                            || dictation_command)
                        {
                            SW.TBrecognized_speech.Text = r.FirstCharToUpper();
                            SW.TBconfidence.Text = c.ToString();
                        }

                        if (current_mode == mode.off)
                        {
                            SW.TBconfidence.Text += "/" + confidence_turning_on;
                            if (c >= confidence_turning_on)
                            {
                                SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                    = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                            }
                            else
                            {
                                SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                    = new SolidColorBrush(Color.FromRgb(232, 4, 4));
                            }
                        }
                        else if (current_mode == mode.command || current_mode == mode.grid || dictation_command)
                        {
                            SW.TBconfidence.Text += "/" + confidence_other_commands;
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
                        }
                        //hide recognized dictation when better dictation is enabled
                        else if (better_dictation == false)
                        {
                            SW.TBconfidence.Text += "/" + confidence_dictation;
                            if (c >= confidence_dictation)
                            {
                                SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                    = new SolidColorBrush(Color.FromRgb(0, 128, 0));
                            }
                            else
                            {
                                SW.TBrecognized_speech.Foreground = SW.TBconfidence.Foreground
                                    = new SolidColorBrush(Color.FromRgb(232, 4, 4));
                            }
                        }
                    }));
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
                    if (current_mode == mode.command && c >= confidence_other_commands
                        && actions != null)
                    {
                        THRcommands = new Thread(() => execute_custom_commands(actions));
                        THRcommands.Start();
                    }
                    //turn off speech recognition
                    else if (r == turn_off && c >= confidence_other_commands &&
                        (current_mode == mode.command 
                        && is_bic_in_general_and_mouse_enabled(bic_type.turn_off)
                        || (current_mode == mode.dictation
                        && is_bic_in_dictation_enabled(bic_type.turn_off))))
                    {
                        if (current_mode == mode.dictation && better_dictation)
                        {
                            toggle_better_dictation();
                        }

                        if (read_recognized_speech) ss.SpeakAsync(r);

                        load_turned_off();
                    }
                    //turn on speech recognition
                    else if (r == turn_on && current_mode == mode.off && c >= confidence_turning_on)
                    {
                        if (read_recognized_speech) ss.SpeakAsync(r);
                        
                        current_mode = last_mode;

                        load_turned_on();
                    }
                    //switch to command mode
                    else if (r == switch_to_command_mode
                        && c >= confidence_other_commands && current_mode == mode.dictation
                        && is_bic_in_dictation_enabled(bic_type.switch_to_command))
                    {
                        if (current_mode == mode.dictation && better_dictation)
                        {
                            toggle_better_dictation();
                        }

                        if (read_recognized_speech) ss.SpeakAsync(r);
                        
                        current_mode = mode.command;
                        
                        load_turned_on();
                    }
                    //switch to dictation mode
                    else if (r == switch_to_dictation_mode
                        && c >= confidence_other_commands && current_mode == mode.command)
                    {
                        if (read_recognized_speech) ss.SpeakAsync(r);
                        
                        current_mode = mode.dictation;
                        
                        load_turned_on();
                    }
                    //show speech recognition window
                    else if (r == show_speech_recognition && current_mode == mode.command
                        && c >= confidence_other_commands)
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
                    }
                    //hide speech recognition window
                    else if (r == hide_speech_recognition && current_mode == mode.command 
                        && c >= confidence_other_commands)
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
                    else if (current_mode == mode.dictation && better_dictation
                        && r == toggle_better_dictation_str
                        && is_bic_in_dictation_enabled(bic_type.toggle_better_dictation))
                    {
                        if (read_recognized_speech) ss.SpeakAsync(r);
                        
                        toggle_better_dictation();
                    }
                    else if (current_mode == mode.dictation && better_dictation
                        && r == start_better_dictation_listening
                        && is_bic_in_dictation_enabled(bic_type.start_better_dictation_listening))
                    {
                        if (read_recognized_speech) ss.SpeakAsync(r);
                        
                        restart_better_dictation();
                    }
                    //dictation text
                    else if (current_mode == mode.dictation && better_dictation == false
                        && c >= confidence_dictation)
                    //&& (r.Length < 6 || r.Substring(0, 6).ToLower() != "press "))
                    {
                        if (read_recognized_speech) ss.SpeakAsync(r);
                        
                        sim.Keyboard.TextEntry(r + " ");
                    }
                    //move mouse grid
                    else if (c >= confidence_other_commands && grid_visible && movable_grid
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

                        THRmouse = new Thread(new ThreadStart(adv_mouse));
                        THRmouse.Start();
                    }
                    //execute a command
                    else if (current_mode == mode.command && c >= confidence_other_commands)
                    {
                        //if (r.Length >= 6 && r.Substring(0, 6).ToLower() != "press ")
                        //    r.Replace("press ", "");
                        THRcommands = new Thread(new ThreadStart(execute_commands));
                        THRcommands.Start();
                    }
                }
                else
                {
                    if (read_recognized_speech)
                    { 
                        ss.SpeakAsync("I'm busy");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error MW009", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            inside_speech_recognized_event = false;
        }        

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
                                enable_grid_gr();
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
                                key_press(list_bic_keys_pressing[i].vkc);
                            else
                                key_press(list_bic_keys_pressing[i].vkc, 1);
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
                            if (current_mode == mode.dictation && better_dictation)
                            {
                                toggle_better_dictation();
                            }

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