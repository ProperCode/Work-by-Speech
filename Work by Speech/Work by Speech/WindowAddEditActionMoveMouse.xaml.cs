using System;
using System.Windows;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditActionMoveMouse : Window
    {
        bool edit = false;
        string prev_action = "";
        int index = -1;

        public WindowAddEditActionMoveMouse(ActionMoveMouse amm, int Index = -1)
        {
            try
            {
                InitializeComponent();

                index = Index;

                Iquestion.ToolTip = "Say \"Get Position\" while in Command mode to set X and Y to " +
                    "current mouse position.";

                if (Middle_Man.last_used_move_position == 0)
                    RBabsolute.IsChecked = true;
                else
                    RBrelative.IsChecked = true;

                if (string.IsNullOrEmpty(amm.action_text) == false)
                {
                    edit = true;
                    prev_action = amm.action_text;

                    this.Title = this.Title.Replace("Add", "Edit");

                    if(amm.absolute)
                        RBabsolute.IsChecked = true;
                    else
                        RBrelative.IsChecked = true;

                    TBx.Text = amm.x.ToString();
                    TBy.Text = amm.y.ToString();

                    if (amm.keys != null)
                    {
                        foreach (string str in amm.keys)
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
                else if (Middle_Man.last_get_position_point.X >= 0
                    && Middle_Man.last_get_position_point.Y >= 0)
                {
                    TBx.Text = Middle_Man.last_get_position_point.X.ToString();
                    TBy.Text = Middle_Man.last_get_position_point.Y.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAMM001", MessageBoxButton.OK, MessageBoxImage.Error);
            }            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                RBabsolute.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAMM002", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAMM003", MessageBoxButton.OK, MessageBoxImage.Error);
            }            
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string x = TBx.Text.Trim();
                string y = TBy.Text.Trim();
                int trash = 0;

                if ((int.TryParse(x, out trash) == false) || trash < 0 || trash > 100000)
                {
                    throw new Exception("X must be a number between 1 and 100000.");
                }
                if ((int.TryParse(y, out trash) == false) || trash < 0 || trash > 100000)
                {
                    throw new Exception("Y must be a number between 1 and 100000.");
                }

                string str = "Move cursor";

                if ((bool)RBabsolute.IsChecked)
                    str += " to: ";
                else if ((bool)RBrelative.IsChecked)
                    str += " by: ";

                str += "(" + x + ", " + y + ")";

                if ((bool)CHBalt.IsChecked || (bool)CHBcontrol.IsChecked || (bool)CHBshift.IsChecked
                    || (bool)CHBwindows.IsChecked)
                {
                    str += " while pressing";

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

                if ((bool)RBabsolute.IsChecked)
                    Middle_Man.last_used_move_position = 0;
                else
                    Middle_Man.last_used_move_position = 1;

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAM004", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAMM005", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAMM006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}