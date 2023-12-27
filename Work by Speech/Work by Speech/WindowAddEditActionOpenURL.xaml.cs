using System;
using System.Windows;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditActionOpenURL : Window
    {
        string prev_action;
        bool edit = false;
        int index = -1;

        public WindowAddEditActionOpenURL(ActionOpenURL aou, int Index = -1)
        {
            try
            {
                InitializeComponent();

                index = Index;

                if (string.IsNullOrEmpty(aou.action_text) == false)
                {
                    edit = true;
                    this.Title = this.Title.Replace("Add", "Edit");
                    prev_action = aou.action_text;

                    TBurl.Text = aou.url;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAOU001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public WindowAddEditActionOpenURL(string name)
        {
            try
            {
                InitializeComponent();

                prev_action = name;
                TBurl.Text = name;
                this.Title = "Edit URL";
                edit = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAOU002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                TBurl.Focus();
                TBurl.CaretIndex = TBurl.Text.Length;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAOU003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string url = TBurl.Text.Trim();

                if (url == "")
                    throw new Exception("URL cannot be empty.");

                string str = "Open URL: " + url;

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

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAOU004", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAOU005", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAOU006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}