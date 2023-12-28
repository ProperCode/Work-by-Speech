using System;
using System.Diagnostics;
using System.Net;
using System.Windows;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditActionScrollMouse : Window
    {
        bool edit = false;
        string prev_action = "";
        int index = -1;

        public WindowAddEditActionScrollMouse(ActionScrollMouse asm, int Index = -1)
        {
            try
            {
                InitializeComponent();

                index = Index;

                RBscroll_up.IsChecked = true;
                TBscrolling_value.Text = Middle_Man.last_used_scrolling_value;

                if (string.IsNullOrEmpty(asm.action_text) == false)
                {
                    edit = true;
                    prev_action = asm.action_text;

                    this.Title = this.Title.Replace("Add", "Edit");

                    if(asm.up)
                        RBscroll_up.IsChecked = true;
                    else
                        RBscroll_down.IsChecked = true;

                    TBscrolling_value.Text = asm.scrolling_value.ToString();

                    if (asm.keys != null)
                    {
                        foreach (string str in asm.keys)
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEASM001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                RBscroll_up.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEASM002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string scrolling_value = TBscrolling_value.Text.Trim();
                int trash = 0;

                if ((int.TryParse(scrolling_value, out trash) == false) || trash <= 0 || trash > 100000)
                {
                    throw new Exception("Scrolling value must be a number between 1 and 100000.");
                }

                string str = "Scroll";

                if ((bool)RBscroll_up.IsChecked)
                    str += " up: ";
                else if ((bool)RBscroll_down.IsChecked)
                    str += " down: ";

                str += scrolling_value;

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

                Middle_Man.last_used_scrolling_value = scrolling_value;

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEASM003", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEASM004", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEASM005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}