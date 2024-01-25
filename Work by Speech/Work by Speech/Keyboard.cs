using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using WindowsInput.Native;

namespace Speech
{
    public partial class MainWindow : Window
    {
        Thread THRkeymaster;

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        private const byte VK_MENU = 0x12;
        private const byte VK_F4 = 0x73;
        private const byte VK_OEM_PLUS = 0xBB; //this presses =, not plus !!!
        //private const byte  = ;
        private const int KEYEVENTF_KEYUP = 0x2;
        private const int KEYEVENTF_KEYDOWN = 0x0;

        void key_press(VirtualKeyCode vkc, bool async, int down_ms = 75)
        {
            if (async)
            {
                THRkeymaster = new Thread(() => key_press(vkc, down_ms));
                THRkeymaster.Start();
            }
            else
                key_press(vkc, down_ms);
        }

        void key_press(VirtualKeyCode vkc, int down_ms = 75)
        {
            //left alt in WindowsInput library is bugged (keyup doesn't work)
            if (vkc == VirtualKeyCode.LMENU)
            {
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                Thread.Sleep(down_ms);
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            }
            //Plus in WindowsInput library is bugged
            else if (vkc == VirtualKeyCode.OEM_PLUS)
            {
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYUP, 0);
                Thread.Sleep(down_ms);
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYDOWN, 0);
            }
            else
            {
                sim.Keyboard.KeyDown(vkc);
                Thread.Sleep(down_ms);
                sim.Keyboard.KeyUp(vkc);
            }
        }

        void key_down(VirtualKeyCode vkc)
        {
            //left alt in WindowsInput library is bugged (keyup doesn't work)
            if (vkc == VirtualKeyCode.LMENU)
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            //Plus in WindowsInput library is bugged
            else if (vkc == VirtualKeyCode.OEM_PLUS)
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYDOWN, 0);
            else if (vkc == VirtualKeyCode.LBUTTON)
                left_down();
            else if (vkc == VirtualKeyCode.RBUTTON)
                right_down();
            else
                sim.Keyboard.KeyDown(vkc);
        }

        void key_up(VirtualKeyCode vkc)
        {
            //left alt in WindowsInput library is bugged (keyup doesn't work)
            if (vkc == VirtualKeyCode.LMENU)
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
            //Plus in WindowsInput library is bugged
            else if (vkc == VirtualKeyCode.OEM_PLUS)
                keybd_event(VK_OEM_PLUS, 0, KEYEVENTF_KEYUP, 0);
            else if (vkc == VirtualKeyCode.LBUTTON)
                left_up();
            else if (vkc == VirtualKeyCode.RBUTTON)
                right_up();
            else
                sim.Keyboard.KeyUp(vkc);
        }

        void release_buttons_and_keys()
        {
            pause_holder = true;

            while (holder_paused == false)
            {
                Thread.Sleep(1);
            }

            List<VirtualKeyCode> list = new List<VirtualKeyCode>(keys_to_hold);

            foreach (VirtualKeyCode vkc in list)
            {
                keys_to_hold.Remove(vkc);

                if (vkc == VirtualKeyCode.LBUTTON)
                    left_up();
                else if (vkc == VirtualKeyCode.RBUTTON)
                    right_up();
                else
                {
                    key_up(vkc);
                }
            }

            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON))
            {
                left_up();
            }

            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON))
            {
                right_up();
            }

            foreach (VirtualKeyCode vkc in (VirtualKeyCode[])Enum.GetValues(typeof(VirtualKeyCode)))
            {
                if (sim.InputDeviceState.IsKeyDown(vkc))
                    key_up(vkc);
            }

            pause_holder = false;
        }

        void release_buttons()
        {
            pause_holder = true;

            while (holder_paused == false)
            {
                Thread.Sleep(1);
            }

            List<VirtualKeyCode> mouse_buttons_list = new List<VirtualKeyCode>()
                { VirtualKeyCode.LBUTTON, VirtualKeyCode.RBUTTON };

            foreach (VirtualKeyCode vkc in mouse_buttons_list)
            {
                if (keys_to_hold.Contains(vkc))
                {
                    keys_to_hold.Remove(vkc);

                    if (vkc == VirtualKeyCode.LBUTTON)
                        left_up();
                    else if (vkc == VirtualKeyCode.RBUTTON)
                        right_up();
                    else
                    {
                        key_up(vkc);
                    }
                }
            }

            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON))
            {
                left_up();
            }

            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON))
            {
                right_up();
            }

            pause_holder = false;
        }

        void add_keys_to_keys_to_hold(List<VirtualKeyCode> list)
        {
            pause_holder = true;

            while (holder_paused == false)
            {
                Thread.Sleep(1);
            }

            int nr = holder_loop_nr;

            foreach (VirtualKeyCode vkc in list)
            {
                keys_to_hold.Add(vkc);
            }

            pause_holder = false;

            while (nr == holder_loop_nr)
            {
                Thread.Sleep(1);
            }
        }

        void remove_keys_from_keys_to_hold(List<VirtualKeyCode> list)
        {
            pause_holder = true;

            while (holder_paused == false)
            {
                Thread.Sleep(1);
            }

            int nr = holder_loop_nr;

            foreach (VirtualKeyCode vkc in list)
            {
                keys_to_hold.Remove(vkc);

                if (vkc == VirtualKeyCode.LBUTTON)
                    left_up();
                else if (vkc == VirtualKeyCode.RBUTTON)
                    right_up();
                else
                {
                    key_up(vkc);
                }
            }

            pause_holder = false;
        }

        bool pause_holder = false;
        bool holder_paused = false;
        int holder_loop_nr = 0;

        //This holding method is better than normal method, because it continues to hold even if
        //user accidently presses held key (and releases it in this way)
        void hold_keys_and_buttons()
        {
            while (true)
            {
                if (current_mode == mode.command)
                {
                    //the only way to hold keys in word and notepad in the same way as
                    //physical keys is to press them every 50ms
                    foreach (VirtualKeyCode vkc in keys_to_hold)
                    {
                        if (vkc == VirtualKeyCode.LBUTTON)
                        {
                            left_down();
                        }
                        else if (vkc == VirtualKeyCode.RBUTTON)
                        {
                            right_down();
                        }
                        else
                            key_down(vkc);
                    }

                    holder_loop_nr++;

                    //Stopwatch sw = new Stopwatch();
                    //sw.Start();

                    for (int i = 0; i < 5; i++)
                    {
                        //Thread.Sleep 5ms takes around 5ms when run from visual studio
                        //but around 15ms when run normally
                        //Thread.Sleep 50ms takes around 50ms when run from visual studio
                        //but around 61ms when run normally
                        //So running app normally adds around 10ms to each Thread.Sleep

                        Thread.Sleep(10);

                        if (pause_holder)
                            break;
                    }

                    //sw.Stop();
                    //MessageBox.Show(sw.ElapsedMilliseconds.ToString() + "ms");
                }
                else
                {
                    Thread.Sleep(100);
                }

                if (pause_holder)
                {
                    holder_paused = true;

                    while (pause_holder)
                    {
                        Thread.Sleep(1);
                    }

                    holder_paused = false;
                }
            }
        }
    }
}