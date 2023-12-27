using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using WindowsInput.Native;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowAddEditActionKeyboard : Window
    {
        List<VirtualKey> keys_to_add = new List<VirtualKey>();
        CollectionView cv;
        bool edit = false;
        string prev_action = "";
        int index = -1;

        public WindowAddEditActionKeyboard(ActionKeyboard ak, int Index = -1)
        {
            InitializeComponent();

            index = Index;
            
            if(string.IsNullOrEmpty(ak.action_text) == false)
            {
                edit = true;
                this.Title = this.Title.Replace("Add", "Edit");

                TBpress_time.Text = ak.time.ToString();
                
                foreach(string str in ak.keys)
                {
                    VirtualKeyCode vkc = VirtualKeyCode.VK_0;

                    foreach(VirtualKey vk in Middle_Man.keys)
                    {
                        if (vk.name == str)
                            vkc = vk.vkc;
                    }

                    keys_to_add.Add(new VirtualKey() { name = str, vkc = vkc });
                }

                prev_action = ak.action_text;
            }
            
            if(ak.time < 1)
            { 
                TBpress_time.Text = Middle_Man.last_used_press_time;
            }

            if(ak.option == 0)
            {
                RBpress.IsChecked = true;
            }
            else if (ak.option == 1)
            {
                RBtoggle.IsChecked = true;
            }
            else if (ak.option == 2)
            {
                RBhold_down.IsChecked = true;
            }
            else if (ak.option == 3)
            {
                RBrelease.IsChecked = true;
            }

            LVnew_keys.ItemsSource = Middle_Man.keys;

            Bdelete.IsEnabled = false;
            Bmove_up.IsEnabled = false;
            Bmove_down.IsEnabled = false;
            Badd.IsEnabled = false;

            MIdelete_keys.IsEnabled = false;
            MImove_up_keys.IsEnabled = false;
            MImove_down_keys.IsEnabled = false;

            LVkeys.ItemsSource = keys_to_add;
            cv = (CollectionView)CollectionViewSource.GetDefaultView(
                LVkeys.ItemsSource);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LVnew_keys.Focus();
        }

        private void RBpress_Checked(object sender, RoutedEventArgs e)
        {
            GBkeys.Header = "Pressed keys";

            TBpress_time.Visibility = Visibility.Visible;
            Lpress_time.Visibility = Visibility.Visible;
        }

        private void RBtoggle_Checked(object sender, RoutedEventArgs e)
        {
            GBkeys.Header = "Toggled keys";

            TBpress_time.Visibility = Visibility.Hidden;
            Lpress_time.Visibility = Visibility.Hidden;
        }

        private void RBhold_down_Checked(object sender, RoutedEventArgs e)
        {
            GBkeys.Header = "Held down keys";

            TBpress_time.Visibility = Visibility.Hidden;
            Lpress_time.Visibility = Visibility.Hidden;
        }

        private void RBrelease_Checked(object sender, RoutedEventArgs e)
        {
            GBkeys.Header = "Released keys";

            TBpress_time.Visibility = Visibility.Hidden;
            Lpress_time.Visibility = Visibility.Hidden;
        }

        private void LVkeys_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            int ind = LVkeys.SelectedIndex;

            if (ind == -1)
            {
                Bdelete.IsEnabled = false;
                Bmove_up.IsEnabled = false;
                Bmove_down.IsEnabled = false;

                MIdelete_keys.IsEnabled = false;
                MImove_up_keys.IsEnabled = false;
                MImove_down_keys.IsEnabled = false;
            }
            else
            {
                bool first_selected = false; //in multi selection
                bool last_selected = false;  //in multi selection

                int ind_first = LVkeys.Items.IndexOf(LVkeys.SelectedItems[0]);
                int ind_last = LVkeys.Items.IndexOf(LVkeys.SelectedItems[LVkeys.SelectedItems.Count - 1]);

                //user may select in both directions
                if (ind_first == 0 || ind_last == 0)
                    first_selected = true;
                else if (ind_first == LVkeys.Items.Count - 1 || ind_last == LVkeys.Items.Count - 1)
                    last_selected = true;

                Bdelete.IsEnabled = true;
                MIdelete_keys.IsEnabled = true;

                if (LVkeys.Items.Count > 1)
                {
                    if (first_selected)
                    {
                        Bmove_up.IsEnabled = false;
                        MImove_up_keys.IsEnabled = false;
                    }
                    else
                    {
                        Bmove_up.IsEnabled = true;
                        MImove_up_keys.IsEnabled = true;
                    }

                    if (last_selected)
                    {
                        Bmove_down.IsEnabled = false;
                        MImove_down_keys.IsEnabled = false;
                    }
                    else
                    {
                        Bmove_down.IsEnabled = true;
                        MImove_down_keys.IsEnabled = true;
                    }
                }
            }
        }
        private void LVnew_keys_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            int ind = LVnew_keys.SelectedIndex;

            if (ind == -1)
            {
                Badd.IsEnabled = false;
            }
            else
            {
                Badd.IsEnabled = true;
            }
        }

        private void Bdelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVkeys.SelectedIndex;

                if (ind != -1)
                {
                    MessageBoxResult dialogResult;

                    if (LVkeys.SelectedItems.Count == 1)
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you " +
                            "want to delete the selected key?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }
                    else
                    {
                        dialogResult = System.Windows.MessageBox.Show("Are you sure you " +
                            "want to delete the selected keys?",
                            "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    }

                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        foreach (VirtualKey key in LVkeys.SelectedItems)
                        {
                            int index = LVkeys.Items.IndexOf(key);

                            keys_to_add.RemoveAt(index);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cv.Refresh();
            }
        }

        private void Bmove_up_Click(object sender, RoutedEventArgs e)
        {
            List<int> indexes = new List<int>();

            try
            {
                if (LVkeys.SelectedIndex != -1)
                {
                    int ind_first = LVkeys.Items.IndexOf(LVkeys.SelectedItems[0]);
                    int ind_last = LVkeys.Items.IndexOf(LVkeys.SelectedItems[LVkeys.SelectedItems.Count - 1]);

                    //when user select from top to bottom
                    if (ind_first <= ind_last)
                    {
                        foreach (VirtualKey key in LVkeys.SelectedItems)
                        {
                            int ind = LVkeys.Items.IndexOf(key);

                            if (ind != 0)
                            {
                                indexes.Add(ind - 1);

                                //new used, because we don't want a reference
                                VirtualKey key2 = new VirtualKey()
                                {
                                    name = keys_to_add[ind].name,
                                    vkc = keys_to_add[ind].vkc
                                };

                                //new used, because we don't want a reference
                                keys_to_add[ind] = new VirtualKey()
                                {
                                    name = keys_to_add[ind - 1].name,
                                    vkc = keys_to_add[ind - 1].vkc
                                };
                                keys_to_add[ind - 1] = key2;
                            }
                        }
                    }
                    //when user select from bottom to top
                    else
                    {
                        for(int i = LVkeys.SelectedItems.Count - 1; i >= 0; i--)
                        {
                            int ind = LVkeys.Items.IndexOf(LVkeys.SelectedItems[i]);

                            if (ind != 0)
                            {
                                indexes.Add(ind - 1);

                                //new used, because we don't want a reference
                                VirtualKey key2 = new VirtualKey()
                                {
                                    name = keys_to_add[ind].name,
                                    vkc = keys_to_add[ind].vkc
                                };

                                //new used, because we don't want a reference
                                keys_to_add[ind] = new VirtualKey()
                                {
                                    name = keys_to_add[ind - 1].name,
                                    vkc = keys_to_add[ind - 1].vkc
                                };
                                keys_to_add[ind - 1] = key2;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cv.Refresh();

                foreach (int ind in indexes)
                    LVkeys.SelectedItems.Add(LVkeys.Items[ind]);
            }
        }

        private void Bmove_down_Click(object sender, RoutedEventArgs e)
        {
            List<int> indexes = new List<int>();

            try
            {
                if (LVkeys.SelectedIndex != -1)
                {
                    int ind_first = LVkeys.Items.IndexOf(LVkeys.SelectedItems[0]);
                    int ind_last = LVkeys.Items.IndexOf(LVkeys.SelectedItems[LVkeys.SelectedItems.Count - 1]);

                    //when user select from bottom to top
                    if (ind_first >= ind_last)
                    {
                        foreach (VirtualKey key in LVkeys.SelectedItems)
                        {
                            int ind = LVkeys.Items.IndexOf(key);

                            if (ind != LVkeys.Items.Count - 1)
                            {
                                indexes.Add(ind + 1);

                                //new used, because we don't want a reference
                                VirtualKey key2 = new VirtualKey()
                                {
                                    name = keys_to_add[ind].name,
                                    vkc = keys_to_add[ind].vkc
                                };

                                //new used, because we don't want a reference
                                keys_to_add[ind] = new VirtualKey()
                                {
                                    name = keys_to_add[ind + 1].name,
                                    vkc = keys_to_add[ind + 1].vkc
                                };
                                keys_to_add[ind + 1] = key2;
                            }
                        }
                    }
                    //when user select from top to bottom
                    else
                    {
                        for (int i = LVkeys.SelectedItems.Count - 1; i >= 0; i--)
                        {
                            int ind = LVkeys.Items.IndexOf(LVkeys.SelectedItems[i]);

                            if (ind != LVkeys.Items.Count - 1)
                            {
                                indexes.Add(ind + 1);

                                //new used, because we don't want a reference
                                VirtualKey key2 = new VirtualKey()
                                {
                                    name = keys_to_add[ind].name,
                                    vkc = keys_to_add[ind].vkc
                                };

                                //new used, because we don't want a reference
                                keys_to_add[ind] = new VirtualKey()
                                {
                                    name = keys_to_add[ind + 1].name,
                                    vkc = keys_to_add[ind + 1].vkc
                                };
                                keys_to_add[ind + 1] = key2;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cv.Refresh();

                foreach (int ind in indexes)
                    LVkeys.SelectedItems.Add(LVkeys.Items[ind]);
            }
        }        

        private void Badd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int ind = LVnew_keys.SelectedIndex;

                if (ind != -1)
                {
                    foreach (VirtualKey key in LVnew_keys.SelectedItems)
                    {
                        bool found = false;

                        foreach(VirtualKey key2 in keys_to_add)
                        {
                            if(key.name == key2.name)
                            {
                                found = true;
                                break;
                            }
                        }

                        if (found == false)
                        {
                            keys_to_add.Add(key);

                            LVkeys.ScrollIntoView(key);
                        }
                        else
                        {
                            throw new Exception(key.name + " is already added.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cv.Refresh();
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string press_time = TBpress_time.Text.Trim();
                int trash = 0;

                if ((bool)RBpress.IsChecked && ((int.TryParse(press_time, out trash) == false) || trash < 0))
                {
                    throw new Exception("Press time must be a number between 1 and 2000000000.");
                }

                if(keys_to_add.Count == 0)
                {
                    throw new Exception("At least one key is required.");
                }

                string str = "";

                if((bool)RBpress.IsChecked)
                    str += "Press: ";
                else if ((bool)RBtoggle.IsChecked)
                    str += "Toggle: ";
                else if ((bool)RBhold_down.IsChecked)
                    str += "Hold down: ";
                else if ((bool)RBrelease.IsChecked)
                    str += "Release: ";

                str += keys_to_add[0].name;

                for(int i=1; i<keys_to_add.Count; i++)
                {
                    str += " + " + keys_to_add[i].name;
                }

                if ((bool)RBpress.IsChecked)
                {
                    str += " for " + press_time + "ms";
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

                Middle_Man.last_used_press_time = press_time;

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK005", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WAEAK006", MessageBoxButton.OK, MessageBoxImage.Error);
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
                else if (e.Key == Key.Delete)
                {
                    Bdelete_Click(null, null);
                }
                else if (e.Key == Key.Up
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control
                    && Bmove_up.IsEnabled)
                {
                    MImove_up_keys_Click(null, null);
                }
                else if (e.Key == Key.Down
                    && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control
                    && Bmove_down.IsEnabled)
                {
                    MImove_down_keys_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK00", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MIdelete_keys_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Bdelete_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK00", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MImove_up_keys_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Bmove_up_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK00", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MImove_down_keys_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Bmove_down_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WAEAK00", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVnew_keys_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (LVnew_keys.SelectedIndex != -1)
                Badd_Click(null, null);
        }
    }
}