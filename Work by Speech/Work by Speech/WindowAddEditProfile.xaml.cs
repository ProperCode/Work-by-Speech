using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditProfile : System.Windows.Window
    {
        bool edit = false;
        string prev_name;

        public WindowAddEditProfile()
        {
            try
            {
                InitializeComponent();

                TBprogram.Text = Middle_Man.any_program_name;
                CHBenabled.IsChecked = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEP001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public WindowAddEditProfile(string name, string program, bool enabled)
        {
            try
            {
                InitializeComponent();

                this.Title = "Edit profile";
                TBname.Text = prev_name = name;
                TBprogram.Text = program;
                CHBenabled.IsChecked = enabled;

                edit = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEP002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                TBname.Focus();
                TBname.CaretIndex = TBname.Text.Length;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEP003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bbrowse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WindowChooseProgram w = new WindowChooseProgram();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEP004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string name = TBname.Text;
                string program = TBprogram.Text;

                if (name == "")
                    throw new Exception("Profile name cannot be empty.");
                else if (program == "")
                    throw new Exception("Associated program name cannot be empty.");
                else if (Middle_Man.contains_illegal_characers_or_names(name))
                {
                    throw new Exception("The following characters cannot be used in the profile name: " +
                        "<, >, :, \\, \", /, |, ?, *.");
                }

                //Add new profile
                if (edit == false)
                {
                    foreach (Profile profile in Middle_Man.profiles)
                    {
                        if (profile.name == name)
                        {
                            throw new Exception("A profile with the same name already exists.");
                        }
                    }

                    Profile p = new Profile()
                    {
                        name = name,
                        program = program,
                        enabled = (bool)CHBenabled.IsChecked,
                        custom_commands = new List<CustomCommand>()
                    };

                    Middle_Man.profiles.Add(p);
                }
                //Update existing profile
                else
                {
                    if (prev_name != name)
                    {
                        foreach (Profile profile in Middle_Man.profiles)
                        {
                            if (profile.name == name)
                            {
                                throw new Exception("A profile with the same name already exists.");
                            }
                        }   
                    }

                    for (int i = 0; i < Middle_Man.profiles.Count; i++)
                    {
                        if (prev_name == Middle_Man.profiles[i].name)
                        {
                            Middle_Man.profiles[i].name = name;
                            Middle_Man.profiles[i].program = program;
                            Middle_Man.profiles[i].enabled = (bool)CHBenabled.IsChecked;
                            break;
                        }
                    }
                }

                Middle_Man.sort_profiles_by_name_asc();
                Middle_Man.save_profiles();

                //delete old profile file if name was changed:
                if (name != prev_name)
                    File.Delete(Path.Combine(Middle_Man.profiles_path, prev_name + ".xml"));

                foreach (System.Windows.Window window in Application.Current.Windows)
                {
                    if (window.GetType() == typeof(MainWindow))
                    {
                        MainWindow w = (MainWindow)window;

                        w.cv_LVprofiles.Refresh();

                        w.LVprofiles.SelectedIndex = Middle_Man.get_profile_ind_by_name(name);

                        w.LVprofiles.ScrollIntoView(w.LVprofiles.SelectedItem);

                        if (edit == false)
                        {
                            //w.Badd_command.Focus(); //it causes problems with pasting commands to a newly
                                                      //aded profiles
                            w.LVprofiles.Focus();
                        }
                        else
                            w.LVprofiles.Focus();
                    }
                }

                Middle_Man.force_updating_both_cc_grammars = true;

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEP005", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEP006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Escape)
                {
                    Bcancel_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEP007", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}