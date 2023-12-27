using System;
using System.IO;
using System.Windows;

namespace Speech
{
    public partial class MainWindow : Window
    {
        const string old_filename_settings = "settings.txt";

        void load_settings_v_1_5_and_older()
        {
            FileStream fs = null;
            StreamReader sr = null;
            string file_path = System.IO.Path.Combine(Middle_Man.saving_folder_path, old_filename_settings);

            try
            {
                if (File.Exists(file_path))
                {
                    fs = new FileStream(file_path, FileMode.Open, FileAccess.Read);
                    sr = new StreamReader(fs);

                    TBconfidence_start.Text = sr.ReadLine();
                    TBconfidence_commands.Text = sr.ReadLine();
                    TBconfidence_dictation.Text = sr.ReadLine();
                    CHBread_recognized_speech.IsChecked = read_recognized_speech
                        = bool.Parse(sr.ReadLine());
                    CHBuse_better_dictation.IsChecked = better_dictation = bool.Parse(sr.ReadLine());

                    CBtype.SelectedIndex = int.Parse(sr.ReadLine());
                    CBlines.SelectedIndex = int.Parse(sr.ReadLine());
                    TBdesired_figures_nr.Text = sr.ReadLine();
                    color_bg_str = sr.ReadLine();
                    color_font_str = sr.ReadLine();
                    TBfont_size.Text = sr.ReadLine();
                    CHBsmart_mousegrid.IsChecked = smart_grid = bool.Parse(sr.ReadLine());

                    CHBstart_with_hidden.IsChecked = bool.Parse(sr.ReadLine());
                    CHBrun_at_startup.IsChecked = bool.Parse(sr.ReadLine());
                    CHBstart_minimized.IsChecked = bool.Parse(sr.ReadLine());
                    CHBminimize_to_tray.IsChecked = bool.Parse(sr.ReadLine());
                    CHBauto_updates.IsChecked = auto_updates = bool.Parse(sr.ReadLine());

                    //Checkboxes Checked and Unchecked events work only after form is loaded
                    //so they have to be called manually in order to load save data properly
                    CHBread_recognized_speech_Checked(new object(), new RoutedEventArgs());
                    CHBuse_better_dictation_Checked(new object(), new RoutedEventArgs());
                    CHBstart_with_hidden_Checked(new object(), new RoutedEventArgs());
                    CHBstart_minimized_Checked(new object(), new RoutedEventArgs());
                    CHBminimize_to_tray_Checked(new object(), new RoutedEventArgs());
                    CHBauto_updates_Checked(new object(), new RoutedEventArgs());
                    CHBsmart_mousegrid_Checked(new object(), new RoutedEventArgs());

                    string line;

                    //added by version 1.3
                    if (sr.EndOfStream == false)
                    {
                        line = sr.ReadLine();
                        if (line != "null" && CBss_voices.Items.Count > 0)
                            CBss_voices.SelectedItem = line;
                    }

                    set_values();

                    sr.Close();
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                loading_error = true;
                MessageBox.Show(ex.Message, "Error OVS001", MessageBoxButton.OK, MessageBoxImage.Error);

                try
                {
                    if (sr != null)
                        sr.Close();
                    if (fs != null)
                        fs.Close();
                }
                catch (Exception ex2) { }
            }
        }
    }
}
