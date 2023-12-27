using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using WindowsInput;
using WindowsInput.Native;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowRecordActions : Window
    {
        class Recorded_Action
        {
            public string action { get; set; }
        }

        Thread THRrecorder;
        bool recording = false;
        bool record_mouse = false; //record mouse movements
        
        public WindowRecordActions()
        {
            try
            {
                InitializeComponent();

                CHBrecord_mouse.IsChecked = Middle_Man.last_used_record_mouse_movements;

                Iquestion.ToolTip = "While in Command mode, say \"Start recording\" to start macro " +
                    "recording and \"Stop recording\" to stop recording.";

                LVactions.SelectionMode = System.Windows.Controls.SelectionMode.Single;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WRA001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void Bstart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Bstart.Content.ToString() == "Start")
                {
                    Bstart.Content = "Stop";
                    this.Title = "Recording Now";

                    THRrecorder = new Thread(new ThreadStart(record));
                    THRrecorder.Start();
                    recording = true;
                }
                else
                {
                    Bstart.Content = "Start";
                    this.Title = "Record Macro";
                    recording = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WRA002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void record()
        {
            try
            {
                InputSimulator sim = new InputSimulator();

                if (LVactions.Items.Count > 0)
                {
                    this.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                    {
                        //data binding is not used, because it causes exceptions when ItemsSource
                        //is updated from another Thread (An ItemsControl is inconsistent with its
                        //items source) and using ObservableCollection seems to be an overkill
                        //https://stackoverflow.com/questions/22524569/an-itemscontrol-is-inconsistent-with-its-items-source-wpf-listbox
                        
                        LVactions.Items.Clear();
                        LVactions.Items.Refresh();
                        Badd.IsEnabled = false;
                    }));
                }

                List<bool> keys_states = new List<bool>();
                
                foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(typeof(VirtualKeyCode)))
                {
                    if (vkc != VirtualKeyCode.LBUTTON
                        && vkc != VirtualKeyCode.MBUTTON
                        && vkc != VirtualKeyCode.RBUTTON
                        && vkc != VirtualKeyCode.XBUTTON1
                        && vkc != VirtualKeyCode.XBUTTON2)
                    {
                        keys_states.Add(sim.InputDeviceState.IsKeyDown(vkc));
                    }
                }

                int prev_mouse_X = System.Windows.Forms.Cursor.Position.X;
                int prev_mouse_Y = System.Windows.Forms.Cursor.Position.Y;

                int curr_mouse_X = prev_mouse_X;
                int curr_mouse_Y = prev_mouse_Y;

                bool lmb_pressed_prev = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON);
                bool rmb_pressed_prev = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON);

                bool lmb_pressed_now = false;
                bool rmb_pressed_now = false;

                List<string> keys_held_down = new List<string>();
                List<string> keys_released = new List<string>();

                bool changes_detected = false;
                bool first_action_added = false;

                Stopwatch sw = new Stopwatch();
                long last_time = 0;

                List<string> prev_actions = new List<string>(); //prev actions (count must be > 0)
                List<string> prev_actions_last_loop = new List<string>(); //prev loop actions (if any)
                List<string> new_actions;

                int loop_nr = 1;
                int sleep_time = 10; //wait time between each loop (in ms)
                                    //10ms is optimal for combining multiple key presses
                                    //and high enough recording accuracy

                while (recording)
                {
                    changes_detected = false;

                    new_actions = new List<string>();

                    //checking mouse position every 50ms is perfect for recording drawing
                    if(record_mouse && loop_nr % 5 == 0)
                    {
                        curr_mouse_X = System.Windows.Forms.Cursor.Position.X;
                        curr_mouse_Y = System.Windows.Forms.Cursor.Position.Y;

                        if (curr_mouse_X != prev_mouse_X || curr_mouse_Y != prev_mouse_Y)
                        {
                            changes_detected = true;

                            new_actions.Add("Move cursor to: (" + curr_mouse_X + ", " 
                                + curr_mouse_Y + ")");

                            prev_mouse_X = curr_mouse_X;
                            prev_mouse_Y = curr_mouse_Y;
                        }
                    }

                    lmb_pressed_now = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON);
                    
                    if (lmb_pressed_now != lmb_pressed_prev)
                    {
                        changes_detected = true;

                        if (lmb_pressed_prev == false)
                        {
                            new_actions.Add("Hold down: LMB (" + System.Windows.Forms.Cursor.Position.X + ", "
                                    + System.Windows.Forms.Cursor.Position.Y + ")");
                        }
                        else
                        {
                            new_actions.Add("Release: LMB (" + System.Windows.Forms.Cursor.Position.X + ", "
                                    + System.Windows.Forms.Cursor.Position.Y + ")");
                        }

                        lmb_pressed_prev = lmb_pressed_now;
                    }

                    rmb_pressed_now = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON);

                    if (rmb_pressed_now != rmb_pressed_prev)
                    {
                        changes_detected = true;

                        if (rmb_pressed_prev == false)
                        {
                            new_actions.Add("Hold down: RMB (" + System.Windows.Forms.Cursor.Position.X + ", "
                                    + System.Windows.Forms.Cursor.Position.Y + ")");
                        }
                        else
                        {
                            new_actions.Add("Release: RMB (" + System.Windows.Forms.Cursor.Position.X + ", "
                                    + System.Windows.Forms.Cursor.Position.Y + ")");
                        }

                        rmb_pressed_prev = rmb_pressed_now;
                    }

                    int ind = 0;
                    bool state;

                    keys_held_down = new List<string>();
                    keys_released = new List<string>();

                    foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(typeof(VirtualKeyCode)))
                    {
                        if (vkc != VirtualKeyCode.LBUTTON
                            && vkc != VirtualKeyCode.MBUTTON
                            && vkc != VirtualKeyCode.RBUTTON
                            && vkc != VirtualKeyCode.XBUTTON1
                            && vkc != VirtualKeyCode.XBUTTON2)
                        {
                            state = sim.InputDeviceState.IsKeyDown(vkc);
                            
                            if (state != keys_states[ind] &&
                                (vkc != VirtualKeyCode.CONTROL 
                                || sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RMENU) == false))
                                //when pressing right alt, control is always recognized as pressed
                                //so control can be only added if right alt was not pressed
                            {
                                string key_name = Middle_Man.get_key_name_by_virtual_key_code(vkc);

                                if (string.IsNullOrEmpty(key_name) == false)
                                {
                                    if (state)
                                        keys_held_down.Add(key_name);
                                    else
                                        keys_released.Add(key_name);
                                }

                                keys_states[ind] = state;
                            }

                            ind++;
                        }
                    }

                    if (keys_held_down.Count > 0)
                    {
                        changes_detected = true;

                        string str = "Hold down: " + keys_held_down[0];

                        for (int i = 1; i < keys_held_down.Count; i++)
                        {
                            str += " + " + keys_held_down[i];
                        }

                        new_actions.Add(str);
                    }

                    if (keys_released.Count > 0)
                    {
                        changes_detected = true;

                        string str = "Release: " + keys_released[0];

                        for (int i = 1; i < keys_released.Count; i++)
                        {
                            str += " + " + keys_released[i];
                        }

                        new_actions.Add(str);
                    }

                    if (changes_detected)
                    {
                        //combining hold and release into click or press
                        if (prev_actions.Count == 1 && new_actions.Count == 1
                            && new_actions[0] == prev_actions[0].Replace("Hold down", "Release"))
                        {
                            if (new_actions[0].Contains("LMB") || new_actions[0].Contains("RMB"))
                            {
                                new_actions[0] = new_actions[0].Replace("Release", "Click")
                                    + " for " + sw.ElapsedMilliseconds + "ms";
                            }
                            else
                            {
                                new_actions[0] = new_actions[0].Replace("Release", "Press")
                                    + " for " + sw.ElapsedMilliseconds + "ms";
                            }

                            if (first_action_added)
                            {
                                LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    LVactions.Items.Add(new Recorded_Action()
                                    { action = "Wait: " + last_time + "ms" });
                                }));
                            }

                            LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                LVactions.Items.Add(new Recorded_Action() { action = new_actions[0] });
                                LVactions.ScrollIntoView(LVactions.Items[LVactions.Items.Count - 1]);
                            }));

                            first_action_added = true;

                            new_actions.Clear();

                            last_time = sw.ElapsedMilliseconds;
                            sw.Restart();
                        }
                        else if (prev_actions_last_loop.Count == 1 && new_actions.Count == 1
                            && (new_actions[0].Contains("Hold") || new_actions[0].Contains("Release"))
                            && (prev_actions[0].Contains("Hold") || prev_actions[0].Contains("Release"))
                            && new_actions[0].Contains("(") == false && prev_actions[0].Contains("(") == false)
                        {
                            //combine hold keys into 1 hold
                            if (prev_actions_last_loop[0].Contains("Hold") && new_actions[0].Contains("Hold"))
                            {
                                string str = prev_actions_last_loop[0].Replace("Hold down: ", "");

                                List<string> keys = new List<string>();

                                if (str.Contains(" + "))
                                {
                                    string[] arr = str.Split(new string[] { " + " },
                                        StringSplitOptions.RemoveEmptyEntries);

                                    foreach (string key in arr)
                                    {
                                        keys.Add(key);
                                    }
                                }
                                else
                                {
                                    keys.Add(str);
                                }

                                str = new_actions[0].Replace("Hold down: ", "");

                                if (str.Contains(" + "))
                                {
                                    string[] arr = str.Split(new string[] { " + " },
                                        StringSplitOptions.RemoveEmptyEntries);

                                    foreach (string key in arr)
                                    {
                                        keys.Add(key);
                                    }
                                }
                                else
                                {
                                    keys.Add(str);
                                }

                                new_actions[0] = "Hold down: " + keys[0];

                                for (int i = 1; i < keys.Count; i++)
                                {
                                    new_actions[0] += " + " + keys[i];
                                }
                            }
                            //combine release keys into 1 release
                            else if (prev_actions_last_loop[0].Contains("Release")
                                && new_actions[0].Contains("Release"))
                            {
                                string str = prev_actions_last_loop[0].Replace("Release: ", "");

                                List<string> keys = new List<string>();

                                if (str.Contains(" + "))
                                {
                                    string[] arr = str.Split(new string[] { " + " },
                                        StringSplitOptions.RemoveEmptyEntries);

                                    foreach (string key in arr)
                                    {
                                        keys.Add(key);
                                    }
                                }
                                else
                                {
                                    keys.Add(str);
                                }

                                str = new_actions[0].Replace("Release: ", "");

                                if (str.Contains(" + "))
                                {
                                    string[] arr = str.Split(new string[] { " + " },
                                        StringSplitOptions.RemoveEmptyEntries);

                                    foreach (string key in arr)
                                    {
                                        keys.Add(key);
                                    }
                                }
                                else
                                {
                                    keys.Add(str);
                                }

                                new_actions[0] = "Release: " + keys[0];

                                for (int i = 1; i < keys.Count; i++)
                                {
                                    new_actions[0] += " + " + keys[i];
                                }

                                LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    LVactions.Items.Add(new Recorded_Action()
                                    { action = "Wait: " + sw.ElapsedMilliseconds + "ms" });

                                    LVactions.Items.Add(new Recorded_Action() { action = new_actions[0] });
                                    LVactions.ScrollIntoView(LVactions.Items[LVactions.Items.Count - 1]);
                                }));

                                first_action_added = true;

                                new_actions.Clear();

                                last_time = sw.ElapsedMilliseconds;
                                sw.Restart();
                            }
                            else
                            {
                                if (first_action_added)
                                {
                                    LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        LVactions.Items.Add(new Recorded_Action()
                                        { action = "Wait: " + sw.ElapsedMilliseconds + "ms" });
                                    }));
                                }

                                LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    LVactions.Items.Add(new Recorded_Action() { action = prev_actions[0] });
                                    LVactions.ScrollIntoView(LVactions.Items[LVactions.Items.Count - 1]);
                                }));

                                first_action_added = true;

                                last_time = sw.ElapsedMilliseconds;
                                sw.Restart();
                            }
                        }
                        else
                        {
                            if (prev_actions.Count > 0)
                            {
                                if (first_action_added)
                                {
                                    LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        LVactions.Items.Add(new Recorded_Action()
                                        { action = "Wait: " + last_time + "ms" });
                                    }));
                                }

                                foreach (string str in prev_actions)
                                {
                                    LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        LVactions.Items.Add(new Recorded_Action() { action = str });
                                    }));

                                    first_action_added = true;
                                }

                                LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    LVactions.ScrollIntoView(LVactions.Items[LVactions.Items.Count - 1]);
                                }));
                            }

                            last_time = sw.ElapsedMilliseconds;
                            sw.Restart();
                        }

                        //copy the list without reference
                        prev_actions = new List<string>(new_actions);
                    }
                    else
                    {
                        bool wait_added = false;

                        int count = 0;

                        for (int i = 0; i < prev_actions.Count; i++)
                        {
                            string str = prev_actions[i];

                            if (str.Contains("Hold") == false || prev_actions.Count > 1)
                            {
                                if (wait_added == false && first_action_added)
                                {
                                    LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                    {
                                        LVactions.Items.Add(new Recorded_Action()
                                        { action = "Wait: " + (last_time + sleep_time) + "ms" });
                                    }));

                                    wait_added = true;
                                }

                                LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                                {
                                    LVactions.Items.Add(new Recorded_Action() { action = str });

                                }));

                                first_action_added = true;

                                prev_actions.RemoveAt(i);
                                i--;
                                count++;

                                last_time = sw.ElapsedMilliseconds;
                                sw.Restart();
                            }
                        }

                        if (count > 0)
                        {
                            LVactions.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                            {
                                LVactions.ScrollIntoView(LVactions.Items[LVactions.Items.Count - 1]);
                            }));
                        }
                    }

                    //copy the list without reference
                    prev_actions_last_loop = new List<string>(new_actions);

                    Thread.Sleep(sleep_time);
                    
                    if(first_action_added)
                    {
                        this.Dispatcher.Invoke(DispatcherPriority.Send, new Action(() =>
                        {
                            if (Badd.IsEnabled == false)
                                Badd.IsEnabled = true;
                        }));
                    }
                    
                    loop_nr++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WRA003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                Bstart.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WRA004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Badd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LVactions.Items.Count > 0)
                {
                    foreach (System.Windows.Window window in Application.Current.Windows)
                    {
                        if (window.GetType() == typeof(WindowAddEditCommand))
                        {
                            WindowAddEditCommand w = (WindowAddEditCommand)window;

                            int insert_index = w.LVactions.Items.Count - 1;

                            foreach (Recorded_Action ra in LVactions.Items)
                            {
                                w.actions.Add(new CC_Action() { action = ra.action });
                            }

                            w.cv_actions.Refresh();

                            w.LVactions.SelectedIndex = w.LVactions.Items.Count - 1;
                            w.LVactions.ScrollIntoView(w.LVactions.SelectedItem);
                        }
                    }

                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WRA005", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WRA006", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WRA007", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CHBrecord_mouse_Checked(object sender, RoutedEventArgs e)
        {
            record_mouse = true;
            Middle_Man.last_used_record_mouse_movements = true;
        }

        private void CHBrecord_mouse_Unchecked(object sender, RoutedEventArgs e)
        {
            record_mouse = false;
            Middle_Man.last_used_record_mouse_movements = false;
        }
    }
}