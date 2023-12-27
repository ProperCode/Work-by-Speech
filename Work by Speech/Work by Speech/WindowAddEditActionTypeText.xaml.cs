using System;
using System.Windows;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditActionTypeText : Window
    {
        bool edit = false;
        string prev_action = "";
        int index = -1;

        public WindowAddEditActionTypeText(ActionTypeText att, int Index = -1)
        {
            try
            {
                InitializeComponent();

                index = Index;

                if (string.IsNullOrEmpty(att.action_text) == false)
                {
                    edit = true;
                    this.Title = this.Title.Replace("Add", "Edit");
                    prev_action = att.action_text;

                    TBtext.Text = att.text;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEATT001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                TBtext.Focus();
                TBtext.CaretIndex = TBtext.Text.Length;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEATT002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string text = TBtext.Text;

                if (text == "")
                    throw new Exception("Text cannot be empty.");

                string str = "Type: " + text;

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
                MessageBox.Show(ex.Message, "Error WAEATT003", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEATT004", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEATT005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}