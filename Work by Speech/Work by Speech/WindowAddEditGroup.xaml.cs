using System;
using System.Windows;
using System.Windows.Input;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditGroup : Window
    {
        string prev_name;
        bool edit = false;

        public WindowAddEditGroup()
        {
            try
            {
                InitializeComponent();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEG001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public WindowAddEditGroup(string name)
        {
            try
            {
                InitializeComponent();

                prev_name = name;
                TBname.Text = name;
                this.Title = "Edit group";
                edit = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEG002", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEG003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string name = TBname.Text;

                if (name == "")
                    throw new Exception("Group name cannot be empty.");

                if (edit == false)
                {
                    foreach (Group group in Middle_Man.groups)
                    {
                        if (name == group.name)
                        {
                            throw new Exception("A group with the same name already exists.");
                        }
                    }

                    Middle_Man.groups.Add(new Group() { name = name });
                }
                else
                {
                    for (int i = 0; i < Middle_Man.groups.Count; i++)
                    {
                        if (prev_name == Middle_Man.groups[i].name)
                        {
                            Middle_Man.groups[i].name = name;
                        }
                    }
                }

                foreach (System.Windows.Window window in Application.Current.Windows)
                {
                    if (window.GetType() == typeof(WindowManageGroups))
                    {
                        WindowManageGroups w = (WindowManageGroups)window;

                        Group group;

                        if(edit)
                            group = (Group)w.LVgroups.SelectedItem;
                        else
                            group = (Group)w.LVgroups.Items[w.LVgroups.Items.Count-1];

                        Middle_Man.sort_groups_by_name_asc();
                        Middle_Man.save_groups();

                        w.cv_LVgroups.Refresh();

                        w.LVgroups.SelectedItem = group;

                        if (w.LVgroups.SelectedIndex != -1)
                            w.LVgroups.ScrollIntoView(w.LVgroups.SelectedItem);
                    }
                }

                foreach (System.Windows.Window window in Application.Current.Windows)
                {
                    if (window.GetType() == typeof(WindowAddEditCommand))
                    {
                        WindowAddEditCommand w = (WindowAddEditCommand)window;

                        string selected = "";

                        if(w.CBgroup.SelectedIndex != -1)
                            selected = w.CBgroup.SelectedItem.ToString();

                        if (selected == prev_name)
                            selected = name;
                        
                        w.CBgroup.Items.Clear();

                        foreach (Group group in Middle_Man.groups)
                        {
                            w.CBgroup.Items.Add(group.name);
                        }

                        if(string.IsNullOrEmpty(selected) == false)
                            w.CBgroup.SelectedItem = selected;
                    }
                }

                foreach(Profile p in Middle_Man.profiles)
                {
                    foreach(CustomCommand cc in p.custom_commands)
                    {
                        if(cc.group == prev_name)
                        {
                            cc.group = name;
                        }
                    }
                }

                Middle_Man.save_profiles();

                foreach (System.Windows.Window window in Application.Current.Windows)
                {
                    if (window.GetType() == typeof(MainWindow))
                    {
                        MainWindow w = (MainWindow)window;
                        w.cv_LVcommands.Refresh();
                    }
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEG004", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEG005", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEG006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}