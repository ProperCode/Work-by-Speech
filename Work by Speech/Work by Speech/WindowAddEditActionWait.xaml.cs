using System;
using System.Windows;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditActionWait : Window
    {
        bool edit = false;
        string prev_action = "";
        int index = -1;

        public WindowAddEditActionWait(ActionWait aw, int Index = -1)
        {
            try
            {
                InitializeComponent();

                index = Index;

                TBtime.Text = Middle_Man.last_used_wait_time;
;
                if (string.IsNullOrEmpty(aw.action_text) == false)
                {
                    edit = true;
                    this.Title = this.Title.Replace("Add", "Edit");
                    prev_action = aw.action_text;

                    TBtime.Text = aw.time.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAW002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                TBtime.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAW003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string time = TBtime.Text.Trim();
                int trash = 0;

                if ((int.TryParse(time, out trash) == false))
                {
                    throw new Exception("Wait time must be a number between 1 and 2000000000.");
                }

                if (trash <= 0)
                    throw new Exception("Wait time must be a number between 1 and 2000000000.");

                string str = "Wait: " + time + "ms";

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

                Middle_Man.last_used_wait_time = time;

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAW004", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAW005", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAW006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}