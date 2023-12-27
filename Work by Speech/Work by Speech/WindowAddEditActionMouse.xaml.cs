using System;
using System.Windows;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditActionMouse : Window
    {
        bool edit = false;
        string prev_action;
        int index = -1;

        public WindowAddEditActionMouse(ActionMouse am, int Index = -1)
        {
            try
            {
                InitializeComponent();

                index = Index;

                Iquestion.ToolTip = "Say \"Get Position\" while in Command mode to set X and Y to " +
                    "current mouse position.";

                TBclick_time.Text = Middle_Man.last_used_click_time;

                if (am.option == 0)
                {
                    RBclick.IsChecked = true;
                }
                else if (am.option == 1)
                {
                    RBtoggle.IsChecked = true;
                }
                else if (am.option == 2)
                {
                    RBhold.IsChecked = true;
                }
                else if (am.option == 3)
                {
                    RBrelease.IsChecked = true;
                }

                if (am.left)
                    RBleft.IsChecked = true;
                else
                    RBright.IsChecked = true;

                CHBchange_position.IsChecked = true;

                if (string.IsNullOrEmpty(am.action_text) == false)
                {
                    edit = true;
                    prev_action = am.action_text;

                    this.Title = this.Title.Replace("Add", "Edit");

                    if (am.x == -1 || am.y == -1)
                        CHBchange_position.IsChecked = false;
                    else
                    {
                        TBx.Text = am.x.ToString();
                        TBy.Text = am.y.ToString();
                    }

                    if (am.keys != null)
                    {
                        foreach (string str in am.keys)
                        {
                            if (str == "Alt")
                                CHBalt.IsChecked = true;
                            if (str == "Control")
                                CHBcontrol.IsChecked = true;
                            if (str == "Shift")
                                CHBshift.IsChecked = true;
                            if (str == "Windows")
                                CHBwindows.IsChecked = true;
                        }
                    }
                }
                else if(Middle_Man.last_get_position_point.X >= 0 
                    && Middle_Man.last_get_position_point.Y >= 0)
                {
                    TBx.Text = Middle_Man.last_get_position_point.X.ToString();
                    TBy.Text = Middle_Man.last_get_position_point.Y.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                RBleft.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RBclick_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                GBadditionally.Header = "Additionally pressed keys";

                Lclick_time.Visibility = Visibility.Visible;
                TBclick_time.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RBhold_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                GBadditionally.Header = "Additionally held down keys";

                Lclick_time.Visibility = Visibility.Hidden;
                TBclick_time.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RBtoggle_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                GBadditionally.Header = "Additionally toggled keys";

                Lclick_time.Visibility = Visibility.Hidden;
                TBclick_time.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RBrelease_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                GBadditionally.Header = "Additionally released keys";

                Lclick_time.Visibility = Visibility.Hidden;
                TBclick_time.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string click_time = TBclick_time.Text.Trim();
                string x = TBx.Text.Trim();
                string y = TBy.Text.Trim();
                int trash = 0;

                if ((bool)RBclick.IsChecked && ((int.TryParse(click_time, out trash) == false) || trash < 0))
                {
                    throw new Exception("Click time must be a number between 1 and 2000000000.");
                }
                if ((bool)CHBchange_position.IsChecked && ((int.TryParse(x, out trash) == false) || trash < 0
                    || trash > 100000))
                {
                    throw new Exception("X must be a number between 1 and 100000.");
                }
                if ((bool)CHBchange_position.IsChecked && ((int.TryParse(y, out trash) == false) || trash < 0
                    || trash > 100000))
                {
                    throw new Exception("Y must be a number between 1 and 100000.");
                }

                string str = "";

                if ((bool)RBclick.IsChecked)
                    str += "Click: ";
                else if ((bool)RBtoggle.IsChecked)
                    str += "Toggle: ";
                else if ((bool)RBhold.IsChecked)
                    str += "Hold down: ";
                else if ((bool)RBrelease.IsChecked)
                    str += "Release: ";

                if ((bool)RBleft.IsChecked)
                {
                    str += "LMB";
                }
                else if ((bool)RBright.IsChecked)
                {
                    str += "RMB";
                }

                if ((bool)CHBchange_position.IsChecked)
                {
                    str += " (" + x + ", " + y + ")";
                }

                if ((bool)CHBalt.IsChecked || (bool)CHBcontrol.IsChecked || (bool)CHBshift.IsChecked
                    || (bool)CHBwindows.IsChecked)
                {
                    if ((bool)RBclick.IsChecked)
                    {
                        str += " while pressing";
                    }
                    else
                    {
                        str += " +";
                    }

                    string str2 = "";

                    if ((bool)CHBalt.IsChecked)
                        str2 += " + Alt";
                    if ((bool)CHBcontrol.IsChecked)
                        str2 += " + Control";
                    if ((bool)CHBshift.IsChecked)
                        str2 += " + Shift";
                    if ((bool)CHBwindows.IsChecked)
                        str2 += " + Windows";

                    str += str2.Remove(0, 2);
                }

                if ((bool)RBclick.IsChecked)
                {
                    str += " for " + click_time + "ms";
                }

                foreach (System.Windows.Window window in Application.Current.Windows)
                {
                    if (window.GetType() == typeof(WindowAddEditCommand))
                    {
                        WindowAddEditCommand w = (WindowAddEditCommand)window;

                        int insert_index = w.LVactions.SelectedIndex + 1;

                        if (edit == false)
                        {
                            if (w.LVactions.SelectedIndex != -1)
                                w.actions.Insert(insert_index, new CC_Action() { action = str });
                            else
                                w.actions.Add(new CC_Action() { action = str });
                        }
                        else
                        {
                            w.actions[index].action = str;
                        }

                        w.cv_actions.Refresh();

                        if (edit == false)
                        {
                            if (insert_index != 0)
                                w.LVactions.SelectedIndex = insert_index;
                            else
                                w.LVactions.SelectedIndex = w.LVactions.Items.Count - 1;
                        }
                        else
                            w.LVactions.SelectedIndex = index;

                        w.LVactions.ScrollIntoView(w.LVactions.SelectedItem);
                    }
                }

                Middle_Man.last_used_click_time = click_time;

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM007", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAM008", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Escape)
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM009", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CHBchange_position_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                Lx.Visibility = Visibility.Visible;
                Ly.Visibility = Visibility.Visible;
                TBx.Visibility = Visibility.Visible;
                TBy.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM010", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CHBchange_position_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                Lx.Visibility = Visibility.Hidden;
                Ly.Visibility = Visibility.Hidden;
                TBx.Visibility = Visibility.Hidden;
                TBy.Visibility = Visibility.Hidden;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM011", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Iquestion_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                MessageBox.Show(Iquestion.ToolTip.ToString(), "Information", MessageBoxButton.OK,
                MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM012", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}