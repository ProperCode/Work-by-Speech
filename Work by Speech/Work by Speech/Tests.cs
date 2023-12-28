using System.IO;
using System.Speech.Recognition;
using System.Threading;
using System.Windows;
using WindowsInput.Native;

namespace Speech
{
    public partial class MainWindow
    {
        short test_mode = 0;
        short test1_on = 1;
        short test2_on = 0;
        short test3_on = 0;
        short test4_on = 0;
        short test5_on = 0;

        void display_grammars_status()
        {
            string grammars_status_str = "";

            foreach (Grammar gr in recognizer.Grammars)
            {
                grammars_status_str += gr.Name + ": " + gr.Enabled + "\r\n";
            }

            MessageBox.Show(grammars_status_str);
        }

        void test1()
        {
            int option = 13;

            if (option == 0)
            {
                WindowAddEditProfile w = new WindowAddEditProfile();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 1)
            {
                WindowChooseProgram w = new WindowChooseProgram();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 2)
            {
                WindowAddEditCommand w = new WindowAddEditCommand("", "", false, 20);
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 3)
            {
                WindowManageGroups w = new WindowManageGroups();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 4)
            {
                WindowAddEditGroup w = new WindowAddEditGroup();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 5)
            {
                WindowAddEditActionKeyboard w = new WindowAddEditActionKeyboard(
                            new ActionKeyboard() { action_text = "", keys = null, option = 0, time = 0 });
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 6)
            {
                WindowAddEditActionMouse w = new WindowAddEditActionMouse(
                            new ActionMouse()
                            {
                                action_text = "",
                                keys = null,
                                left = true,
                                option = 0,
                                time = 0,
                                x = -1,
                                y = -1
                            });
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 7)
            {
                WindowAddEditActionMoveMouse w = new WindowAddEditActionMoveMouse(
                            new ActionMoveMouse()
                            {
                                action_text = "",
                                keys = null,
                                absolute = true,
                                x = -1,
                                y = -1
                            });
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 8)
            {
                WindowAddEditActionScrollMouse w = new WindowAddEditActionScrollMouse(
                            new ActionScrollMouse()
                            {
                                action_text = "",
                                keys = null,
                                up = true,
                                scrolling_value = 40
                            });
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 9)
            {
                WindowAddEditActionOpenURL w = new WindowAddEditActionOpenURL(
                            new ActionOpenURL()
                            {
                                action_text = "",
                                url = ""
                            });
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 10)
            {
                WindowAddEditActionReadText w = new WindowAddEditActionReadText(
                            new ActionReadAloud()
                            {
                                action_text = "",
                                text = ""
                            });
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 11)
            {
                WindowAddEditActionTypeText w = new WindowAddEditActionTypeText(
                            new ActionTypeText()
                            {
                                action_text = "",
                                text = ""
                            });
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 12)
            {
                WindowAddEditActionWait w = new WindowAddEditActionWait(
                            new ActionWait()
                            {
                                action_text = "",
                                time = 1000
                            });
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
            else if (option == 13)
            {
                WindowRecordActions w = new WindowRecordActions();
                w.Owner = Application.Current.MainWindow;
                w.ShowInTaskbar = false;
                w.ShowDialog();
            }
        }

        void test2()
        {
            int option = 4;

            //-----------SYMBOLS TYPING TEST START-----------
            if (option == 1)
            {
                Thread.Sleep(5000);
                sim.Keyboard.TextEntry("'V\"V[V-V]+V=V:V“V>V<V~V@V!V?V#V£V$V%V^V(V)V_V{V}V|V™V¾V¼V½V&V//" +
                    "V`V<V>V±V«V»V×V÷V¢V¥V§V©V®V°V¶V...VƑVµ");

                Thread.Sleep(50000);
            }
            //-----------SYMBOLS TYPING TEST END-------------

            //-----------SINGLE KEYS PRESSING TESTS START-----------
            else if (option == 2)
            {
                Thread.Sleep(10000);

                key_press(VirtualKeyCode.OEM_4);

                //LMENU is bugged(keyup doesn't work for it)

                sim.Keyboard.KeyDown(VirtualKeyCode.LMENU);
                sim.Keyboard.KeyDown(VirtualKeyCode.F4);
                Thread.Sleep(75);
                sim.Keyboard.KeyUp(VirtualKeyCode.F4);
                sim.Keyboard.KeyUp(VirtualKeyCode.LMENU);

                //this works now:
                key_press(VirtualKeyCode.LMENU);

                sim.Keyboard.KeyDown(VirtualKeyCode.RMENU);
                sim.Keyboard.KeyDown(VirtualKeyCode.VK_S);
                Thread.Sleep(75);
                sim.Keyboard.KeyUp(VirtualKeyCode.VK_S);
                sim.Keyboard.KeyUp(VirtualKeyCode.RMENU);

                Thread.Sleep(50000);
            }
            //-----------SINGLE KEYS PRESSING TESTS END-------------

            //-----------KEYS PRESSING TEST START-----------
            else if (option == 3)
            {
                Thread.Sleep(10000);

                create_bic_lists();

                foreach (BuiltInCommand s in list_bic_keys_pressing)
                {
                    if (s.vkc != VirtualKeyCode.SPACE
                        && s.name.ToLower().Contains("function") == false
                        && s.vkc != VirtualKeyCode.LWIN
                        && s.vkc != VirtualKeyCode.VOLUME_DOWN
                        && s.vkc != VirtualKeyCode.VOLUME_MUTE
                        && s.vkc != VirtualKeyCode.VOLUME_UP
                        && s.vkc != VirtualKeyCode.MENU
                        && s.vkc != VirtualKeyCode.LMENU
                        && s.vkc != VirtualKeyCode.RMENU
                        && s.vkc != VirtualKeyCode.CAPITAL)
                    {
                        sim.Keyboard.TextEntry(s.name + ": ");
                        Thread.Sleep(100);
                        key_press(s.vkc);
                        Thread.Sleep(100);
                        key_down(VirtualKeyCode.LSHIFT);
                        key_press(VirtualKeyCode.RETURN);
                        key_up(VirtualKeyCode.LSHIFT);
                        Thread.Sleep(100);
                    }
                }

                Thread.Sleep(50000);
            }
            else if (option == 4)
            {
                Thread.Sleep(10000);

                foreach (VirtualKey vk in Middle_Man.keys)
                {
                    if (vk.vkc != VirtualKeyCode.SPACE
                        && vk.name.StartsWith("F") == false
                        && vk.vkc != VirtualKeyCode.LWIN
                        && vk.vkc != VirtualKeyCode.VOLUME_DOWN
                        && vk.vkc != VirtualKeyCode.VOLUME_MUTE
                        && vk.vkc != VirtualKeyCode.VOLUME_UP
                        && vk.vkc != VirtualKeyCode.MENU
                        && vk.vkc != VirtualKeyCode.LMENU
                        && vk.vkc != VirtualKeyCode.RMENU
                        && vk.vkc != VirtualKeyCode.CAPITAL)
                    {
                        sim.Keyboard.TextEntry(vk.name + ": ");
                        Thread.Sleep(100);
                        key_press(vk.vkc);
                        Thread.Sleep(100);
                        key_down(VirtualKeyCode.LSHIFT);
                        key_press(VirtualKeyCode.RETURN);
                        key_up(VirtualKeyCode.LSHIFT);
                        Thread.Sleep(100);
                    }
                }

                Thread.Sleep(50000);
            }
            //-----------KEYS PRESSING TEST END-------------
        }

        void test3()
        {
            FileStream fs = null;
            StreamWriter sw = null;

            fs = new FileStream(System.IO.Path.Combine(new string[] {
                                        Middle_Man.app_folder_path, "test.txt" }),
                    FileMode.Create, FileAccess.Write);
            sw = new StreamWriter(fs);

            Choices ch = new Choices();

            foreach (Installed_App app in installed_apps)
            {
                foreach (string name2 in app.names)
                {
                    if (name2 != null && name2 != "")
                    {
                        ch.Add(new string[] { name2 });
                        sw.WriteLine(name2);
                    }
                }
            }

            sw.Close();
            fs.Close();
        }

        void test4()
        {
            bool state1, state2, state3, state4, state5;

            state1 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON);

            right_down();
            //key_down(VirtualKeyCode.RBUTTON); //doesn't work

            state2 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON);

            Thread.Sleep(1);

            state3 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON);

            right_up();
            //key_up(VirtualKeyCode.RBUTTON); //doesn't work

            state4 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON);

            Thread.Sleep(1);

            state5 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON);

            MessageBox.Show(
                "Is RMB down: " + state1.ToString() + " - before pressing it\r\n"
                + "Is RMB down: " + state2.ToString() + " - immediately after pressing it\r\n"
                + "Is RMB down: " + state3.ToString() + " - 1ms after pressing it\r\n"
                + "Is RMB down: " + state4.ToString() + " - immediately after releasing it\r\n"
                + "Is RMB down: " + state5.ToString() + " - 1ms after releasing it"
                );
        }

        void test5()
        {
            bool state1, state2, state3, state4, state5;

            state1 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON);

            left_down();
            //key_down(VirtualKeyCode.LBUTTON); //doesn't work

            state2 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON);

            Thread.Sleep(1);

            state3 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON);

            left_up();
            //key_up(VirtualKeyCode.LBUTTON); //doesn't work

            state4 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON);

            Thread.Sleep(1);

            state5 = sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON);

            MessageBox.Show(
                "Is LMB down: " + state1.ToString() + " - before pressing it\r\n"
                + "Is LMB down: " + state2.ToString() + " - immediately after pressing it\r\n"
                + "Is LMB down: " + state3.ToString() + " - 1ms after pressing it\r\n"
                + "Is LMB down: " + state4.ToString() + " - after releasing it\r\n"
                + "Is LMB down: " + state5.ToString() + " - 1ms after releasing it"
                );
        }
    }
}
