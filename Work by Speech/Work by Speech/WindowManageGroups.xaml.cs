using System;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowManageGroups : Window
    {
        public CollectionView cv_LVgroups;

        public WindowManageGroups()
        {
            try
            {
                InitializeComponent();

                Bedit.IsEnabled = false;
                Bdelete.IsEnabled = false;

                MIedit.IsEnabled = false;
                MIdelete.IsEnabled = false;

                LVgroups.ItemsSource = Middle_Man.groups;

                cv_LVgroups = (CollectionView)CollectionViewSource.GetDefaultView(
                        LVgroups.ItemsSource);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WMG001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                LVgroups.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WMG002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVgroups_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                int ind = LVgroups.SelectedIndex;

                if (ind != -1)
                {
                    Bedit.IsEnabled = true;
                    Bdelete.IsEnabled = true;

                    MIedit.IsEnabled = true;
                    MIdelete.IsEnabled = true;
                }
                else
                {
                    Bedit.IsEnabled = false;
                    Bdelete.IsEnabled = false;

                    MIedit.IsEnabled = false;
                    MIdelete.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WMG003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Badd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WindowAddEditGroup w = new WindowAddEditGroup();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WMG004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bedit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVgroups.SelectedIndex;

                if (ind != -1)
                {
                    foreach (Group g in LVgroups.SelectedItems)
                    {
                        for (int i = 0; i < Middle_Man.groups.Count; i++)
                        {
                            if (g.name == Middle_Man.general_group_name)
                            {
                                throw new Exception("General group cannot be edited.");
                            }
                        }
                    }

                    WindowAddEditGroup w = new WindowAddEditGroup(
                        ((Group)LVgroups.SelectedItems[0]).name);
                    w.Owner = Application.Current.MainWindow;
                    w.ShowInTaskbar = false;
                    w.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WMG005", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bdelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVgroups.SelectedIndex;

                if (ind != -1)
                {
                    MessageBoxResult dialogResult;

                    if (LVgroups.SelectedItems.Count == 1)
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you want" +
                            " to permanently delete the selected group?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }
                    else
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you want" +
                            " to permanently delete the selected groups?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }

                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        foreach (Group g in LVgroups.SelectedItems)
                        {
                            for (int i = 0; i < Middle_Man.groups.Count; i++)
                            {
                                if (g.name == Middle_Man.general_group_name)
                                {
                                    throw new Exception("General group cannot be deleted.");
                                }
                                else if (g.name == Middle_Man.groups[i].name)
                                {
                                    Middle_Man.groups.RemoveAt(i);
                                    break;
                                }
                            }
                        }

                        Middle_Man.save_groups();

                        foreach (System.Windows.Window window in Application.Current.Windows)
                        {
                            if (window.GetType() == typeof(WindowAddEditCommand))
                            {
                                WindowAddEditCommand w = (WindowAddEditCommand)window;

                                string selected = "";

                                if (w.CBgroup.SelectedIndex != -1)
                                    selected = w.CBgroup.SelectedItem.ToString();

                                w.CBgroup.Items.Clear();

                                foreach (Group group in Middle_Man.groups)
                                {
                                    w.CBgroup.Items.Add(group.name);
                                }

                                if (string.IsNullOrEmpty(selected) == false)
                                    w.CBgroup.SelectedItem = selected;
                            }
                        }

                        foreach (System.Windows.Window window in Application.Current.Windows)
                        {
                            if (window.GetType() == typeof(MainWindow))
                            {
                                MainWindow w = (MainWindow)window;
                                w.cv_LVcommands.Refresh();
                            }
                        }
                    }

                    cv_LVgroups.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WMG006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cv_LVgroups.Refresh();
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WMG007", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WMG008", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVgroups_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Delete)
                {
                    Bdelete_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WMG009", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVgroups_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (LVgroups.SelectedIndex != -1)
                Bedit_Click(null, null);
        }

        private void MIedit_Click(object sender, RoutedEventArgs e)
        {
            if (LVgroups.SelectedIndex != -1)
                Bedit_Click(null, null);
        }

        private void MIdelete_Click(object sender, RoutedEventArgs e)
        {
            if (LVgroups.SelectedIndex != -1)
                Bdelete_Click(null, null);
        }
    }
}