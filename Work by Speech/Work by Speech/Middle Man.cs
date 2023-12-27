using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml;
using WindowsInput.Native;

namespace Speech
{
    static class Middle_Man
    {
        public static string prog_name = "Work by Speech";
        
        public static string groups_filename = "groups.xml";
        public static string bic_filename = "built-in commands.xml";
        public static string profiles_foldername = "profiles";

        public static string general_group_name = "General";
        public static string any_program_name = "Any";

        public static string users_directory_path
            = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        //full path is necessary if run at startup is used (running at startup uses different current
        //directory
        public static string app_folder_path = System.IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location);
        public static string saving_folder_path;
        public static string profiles_path;

        public static string registry_path_easy = "SOFTWARE\\Microsoft\\Work by Speech"; //easy to find
        
        public static string registry_key_first_run = "already run";
        public static string registry_key_last_import_path = "last import profiles path";
        public static string registry_key_last_export_path = "last export profiles path";
        public static string registry_key_last_open_program_path = "last open file/program path";
        public static string registry_key_last_open_sound_file_path = "last open sound file path";

        public static string url_homepage = "github.com/ProperCode/Work-by-Speech";
        public static string url_homepage_full = "https://github.com/ProperCode/Work-by-Speech";
        public static string url_latest_version = "https://raw.githubusercontent.com/ProperCode/Work-by-Speech/main/other/latest_version.txt";
        public static string url_download = "https://github.com/ProperCode/Work-by-Speech";

        public static List<Profile> profiles = new List<Profile>();
        public static List<Group> groups = new List<Group>() { new Group() { name = general_group_name } };
        public static List<CC_Action> copied_actions = new List<CC_Action>();
        public static List<CustomCommand> copied_commands = new List<CustomCommand>();

        public static string last_used_max_executions = "1";
        public static string last_used_press_time = "1";
        public static string last_used_click_time = "1";
        public static short last_used_move_position = 0;
        public static string last_used_scrolling_value = "4";
        public static string last_used_wait_time = "1000";
        public static bool last_used_record_mouse_movements = false;
        public static Point last_get_position_point = new Point(-1, -1); //acquired by saying "Get Position"
                                                                         //while in command mode

        public static bool force_updating_both_cc_grammars = false; //cc = custom commands
        
        public static List<CC_Action> action_types = new List<CC_Action>()
        {
            new CC_Action() { action = "Press key(s)" },
            new CC_Action() { action = "Toggle key(s)" },
            new CC_Action() { action = "Hold down key(s)" },
            new CC_Action() { action = "Release key(s)" },
            new CC_Action() { action = "Release all keys and buttons" },
            new CC_Action() { action = "Click mouse button" },
            new CC_Action() { action = "Toggle mouse button" },
            new CC_Action() { action = "Hold down mouse button" },
            new CC_Action() { action = "Release mouse button" },
            new CC_Action() { action = "Move mouse cursor" },
            new CC_Action() { action = "Scroll mouse wheel" },
            new CC_Action() { action = "Open file/program" },
            new CC_Action() { action = "Open URL" },
            new CC_Action() { action = "Play sound" },
            new CC_Action() { action = "Read aloud text" },
            new CC_Action() { action = "Type text" },
            new CC_Action() { action = "Wait" }
        };

        public static List<VirtualKey> keys = new List<VirtualKey>()
        {
            new VirtualKey() { name = "Control", vkc = VirtualKeyCode.CONTROL },
            new VirtualKey() { name = "Alt", vkc = VirtualKeyCode.LMENU },
            new VirtualKey() { name = "Right Alt", vkc = VirtualKeyCode.RMENU },
            new VirtualKey() { name = "Shift", vkc = VirtualKeyCode.SHIFT },
            new VirtualKey() { name = "Windows", vkc = VirtualKeyCode.LWIN },
            new VirtualKey() { name = "A", vkc = VirtualKeyCode.VK_A },
            new VirtualKey() { name = "B", vkc = VirtualKeyCode.VK_B },
            new VirtualKey() { name = "C", vkc = VirtualKeyCode.VK_C },
            new VirtualKey() { name = "D", vkc = VirtualKeyCode.VK_D },
            new VirtualKey() { name = "E", vkc = VirtualKeyCode.VK_E },
            new VirtualKey() { name = "F", vkc = VirtualKeyCode.VK_F },
            new VirtualKey() { name = "G", vkc = VirtualKeyCode.VK_G },
            new VirtualKey() { name = "H", vkc = VirtualKeyCode.VK_H },
            new VirtualKey() { name = "I", vkc = VirtualKeyCode.VK_I },
            new VirtualKey() { name = "J", vkc = VirtualKeyCode.VK_J },
            new VirtualKey() { name = "K", vkc = VirtualKeyCode.VK_K },
            new VirtualKey() { name = "L", vkc = VirtualKeyCode.VK_L },
            new VirtualKey() { name = "M", vkc = VirtualKeyCode.VK_M },
            new VirtualKey() { name = "N", vkc = VirtualKeyCode.VK_N },
            new VirtualKey() { name = "O", vkc = VirtualKeyCode.VK_O },
            new VirtualKey() { name = "P", vkc = VirtualKeyCode.VK_P },
            new VirtualKey() { name = "Q", vkc = VirtualKeyCode.VK_Q },
            new VirtualKey() { name = "R", vkc = VirtualKeyCode.VK_R },
            new VirtualKey() { name = "S", vkc = VirtualKeyCode.VK_S },
            new VirtualKey() { name = "T", vkc = VirtualKeyCode.VK_T },
            new VirtualKey() { name = "U", vkc = VirtualKeyCode.VK_U },
            new VirtualKey() { name = "V", vkc = VirtualKeyCode.VK_V },
            new VirtualKey() { name = "W", vkc = VirtualKeyCode.VK_W },
            new VirtualKey() { name = "X", vkc = VirtualKeyCode.VK_X },
            new VirtualKey() { name = "Y", vkc = VirtualKeyCode.VK_Y },
            new VirtualKey() { name = "Z", vkc = VirtualKeyCode.VK_Z },
            new VirtualKey() { name = ",", vkc = VirtualKeyCode.OEM_COMMA },
            new VirtualKey() { name = ".", vkc = VirtualKeyCode.OEM_PERIOD },
            new VirtualKey() { name = "/", vkc = VirtualKeyCode.DIVIDE },
            new VirtualKey() { name = "-", vkc = VirtualKeyCode.OEM_MINUS },
            new VirtualKey() { name = "*", vkc = VirtualKeyCode.MULTIPLY },
            new VirtualKey() { name = "; (US keyboard layout)", vkc = VirtualKeyCode.OEM_1 },
            new VirtualKey() { name = "' (US keyboard layout)", vkc = VirtualKeyCode.OEM_7 },
            new VirtualKey() { name = "\\ (US keyboard layout)", vkc = VirtualKeyCode.OEM_5 },
            new VirtualKey() { name = "[ (US keyboard layout)", vkc = VirtualKeyCode.OEM_4 },
            new VirtualKey() { name = "] (US keyboard layout)", vkc = VirtualKeyCode.OEM_6 },
            new VirtualKey() { name = "0", vkc = VirtualKeyCode.VK_0 },
            new VirtualKey() { name = "1", vkc = VirtualKeyCode.VK_1 },
            new VirtualKey() { name = "2", vkc = VirtualKeyCode.VK_2 },
            new VirtualKey() { name = "3", vkc = VirtualKeyCode.VK_3 },
            new VirtualKey() { name = "4", vkc = VirtualKeyCode.VK_4 },
            new VirtualKey() { name = "5", vkc = VirtualKeyCode.VK_5 },
            new VirtualKey() { name = "6", vkc = VirtualKeyCode.VK_6 },
            new VirtualKey() { name = "7", vkc = VirtualKeyCode.VK_7 },
            new VirtualKey() { name = "8", vkc = VirtualKeyCode.VK_8 },
            new VirtualKey() { name = "9", vkc = VirtualKeyCode.VK_9 },
            new VirtualKey() { name = "F1", vkc = VirtualKeyCode.F1 },
            new VirtualKey() { name = "F2", vkc = VirtualKeyCode.F2 },
            new VirtualKey() { name = "F3", vkc = VirtualKeyCode.F3 },
            new VirtualKey() { name = "F4", vkc = VirtualKeyCode.F4 },
            new VirtualKey() { name = "F5", vkc = VirtualKeyCode.F5 },
            new VirtualKey() { name = "F6", vkc = VirtualKeyCode.F6 },
            new VirtualKey() { name = "F7", vkc = VirtualKeyCode.F7 },
            new VirtualKey() { name = "F8", vkc = VirtualKeyCode.F8 },
            new VirtualKey() { name = "F9", vkc = VirtualKeyCode.F9 },
            new VirtualKey() { name = "F10", vkc = VirtualKeyCode.F10 },
            new VirtualKey() { name = "F11", vkc = VirtualKeyCode.F11 },
            new VirtualKey() { name = "F12", vkc = VirtualKeyCode.F12 },
            new VirtualKey() { name = "Backspace", vkc = VirtualKeyCode.BACK },
            new VirtualKey() { name = "Browser back", vkc = VirtualKeyCode.BROWSER_BACK },
            new VirtualKey() { name = "Browser next", vkc = VirtualKeyCode.BROWSER_FORWARD },
            new VirtualKey() { name = "Browser search", vkc = VirtualKeyCode.BROWSER_SEARCH },
            new VirtualKey() { name = "Browser refresh", vkc = VirtualKeyCode.BROWSER_REFRESH },
            new VirtualKey() { name = "Caps lock", vkc = VirtualKeyCode.CAPITAL },
            new VirtualKey() { name = "Delete", vkc = VirtualKeyCode.DELETE },
            new VirtualKey() { name = "Down", vkc = VirtualKeyCode.DOWN },
            new VirtualKey() { name = "End", vkc = VirtualKeyCode.END },
            new VirtualKey() { name = "Enter", vkc = VirtualKeyCode.RETURN },
            new VirtualKey() { name = "Escape", vkc = VirtualKeyCode.ESCAPE },
            new VirtualKey() { name = "Home", vkc = VirtualKeyCode.HOME },
            new VirtualKey() { name = "Insert", vkc = VirtualKeyCode.INSERT },
            new VirtualKey() { name = "Left", vkc = VirtualKeyCode.LEFT },
            new VirtualKey() { name = "Media play pause", vkc = VirtualKeyCode.MEDIA_PLAY_PAUSE },
            new VirtualKey() { name = "Media stop", vkc = VirtualKeyCode.MEDIA_STOP },
            new VirtualKey() { name = "Media prev track", vkc = VirtualKeyCode.MEDIA_PREV_TRACK },
            new VirtualKey() { name = "Media next track", vkc = VirtualKeyCode.MEDIA_NEXT_TRACK },            
            new VirtualKey() { name = "Page down", vkc = VirtualKeyCode.NEXT},
            new VirtualKey() { name = "Page up", vkc = VirtualKeyCode.PRIOR },
            new VirtualKey() { name = "Print screen", vkc = VirtualKeyCode.SNAPSHOT },
            new VirtualKey() { name = "Right", vkc = VirtualKeyCode.RIGHT },
            new VirtualKey() { name = "Space", vkc = VirtualKeyCode.SPACE },
            new VirtualKey() { name = "Tab", vkc = VirtualKeyCode.TAB },
            new VirtualKey() { name = "Up", vkc = VirtualKeyCode.UP },
            new VirtualKey() { name = "Volume down", vkc = VirtualKeyCode.VOLUME_DOWN },
            new VirtualKey() { name = "Volume up", vkc = VirtualKeyCode.VOLUME_UP },
            new VirtualKey() { name = "Volume mute", vkc = VirtualKeyCode.VOLUME_MUTE }
        };

        public static string ByteArrayToHexString(byte[] bytes)
        {
            return string.Join(string.Empty, Array.ConvertAll(bytes, b => b.ToString("X2")));
        }

        public static void save_profiles(List<Profile> chosen_profiles = null, string path = null)
        {
            if (chosen_profiles == null)
                chosen_profiles = profiles;

            string saving_path;

            foreach (Profile p in chosen_profiles)
            {
                XmlDocument xml_doc = new XmlDocument();

                XmlNode root_node = xml_doc.CreateElement("profile");

                XmlAttribute attribute = xml_doc.CreateAttribute("version");
                attribute.Value = "1";
                root_node.Attributes.Append(attribute);

                xml_doc.AppendChild(root_node);

                XmlNode name_node = xml_doc.CreateElement("name");
                name_node.InnerText = p.name;
                root_node.AppendChild(name_node);

                XmlNode program_node = xml_doc.CreateElement("program");
                program_node.InnerText = p.program;
                root_node.AppendChild(program_node);

                XmlNode enabled_node = xml_doc.CreateElement("enabled");
                enabled_node.InnerText = p.enabled.ToString();
                root_node.AppendChild(enabled_node);

                XmlNode ccs_node = xml_doc.CreateElement("custom_commands");
                root_node.AppendChild(ccs_node);
                
                foreach (CustomCommand cc in p.custom_commands)
                {
                    XmlNode cc_node = xml_doc.CreateElement("custom_command");
                    ccs_node.AppendChild(cc_node);

                    XmlNode c_node = xml_doc.CreateElement("name");
                    c_node.InnerText = cc.name;
                    cc_node.AppendChild(c_node);

                    XmlNode description_node = xml_doc.CreateElement("description");
                    description_node.InnerText = cc.description;
                    cc_node.AppendChild(description_node);

                    XmlNode group_node = xml_doc.CreateElement("group");
                    group_node.InnerText = cc.group.ToString();
                    cc_node.AppendChild(group_node);

                    XmlNode max_executions_node = xml_doc.CreateElement("max_executions");
                    max_executions_node.InnerText = cc.max_executions.ToString();
                    cc_node.AppendChild(max_executions_node);

                    enabled_node = xml_doc.CreateElement("enabled");
                    enabled_node.InnerText = cc.enabled.ToString();
                    cc_node.AppendChild(enabled_node);

                    XmlNode actions_node = xml_doc.CreateElement("actions");
                    cc_node.AppendChild(actions_node);

                    foreach (CC_Action action in cc.actions)
                    {
                        XmlNode action_node = xml_doc.CreateElement("action");
                        
                        action_node.InnerText = action.action;

                        actions_node.AppendChild(action_node);
                    }
                }

                if (Directory.Exists(profiles_path) == false)
                {
                    Directory.CreateDirectory(profiles_path);
                }

                if (path == null)
                    saving_path = Path.Combine(profiles_path, p.name + ".xml");
                else
                    saving_path = Path.Combine(path, p.name + ".xml");

                xml_doc.Save(saving_path);
            }
        }

        public static void save_groups()
        {
            XmlDocument xml_doc = new XmlDocument();

            XmlNode root_node = xml_doc.CreateElement("groups");

            XmlAttribute attribute = xml_doc.CreateAttribute("version");
            attribute.Value = "1";
            root_node.Attributes.Append(attribute);

            xml_doc.AppendChild(root_node);

            foreach (Group g in groups)
            {
                XmlNode g_node = xml_doc.CreateElement("group");

                g_node.InnerText = g.name;

                root_node.AppendChild(g_node);
            }

            if (Directory.Exists(saving_folder_path) == false)
            {
                Directory.CreateDirectory(saving_folder_path);
            }

            xml_doc.Save(Path.Combine(saving_folder_path, groups_filename));
        }

        public static void load_groups()
        {
            if (File.Exists(Path.Combine(saving_folder_path, groups_filename)))
            {
                //groups = new ObservableCollection<object>();
                groups = new List<Group>();

                XmlDocument xml_doc = new XmlDocument();
                xml_doc.Load(Path.Combine(saving_folder_path, groups_filename));

                XmlNodeList groups_tag = xml_doc.SelectNodes("//groups");

                int version = -1;
                bool parsing_v = false;

                if (groups_tag[0].Attributes["version"] != null)
                    parsing_v = int.TryParse(groups_tag[0].Attributes["version"].Value, out version);

                //Work by Speech v. 1.5 and earlier had no version attribute (so we treat groups file made by
                //these versions as as version 1 of xml saving method for groups)

                if (version == -1 || (parsing_v && version == 1))
                {
                    XmlNodeList nodes = xml_doc.SelectNodes("//groups/group");

                    foreach (XmlNode node in nodes)
                    {
                        Group g = new Group();

                        g.name = node.InnerText;

                        groups.Add(g);
                    }

                    bool found = false;

                    foreach (Group g in groups)
                    {
                        if (g.name == general_group_name)
                        {
                            found = true;
                            break;
                        }
                    }

                    if (found == false)
                    {
                        groups.Add(new Group() { name = general_group_name });
                    }

                    Middle_Man.sort_groups_by_name_asc();
                }
            }
        }

        public static void sort_profiles_by_name_asc()
        {
            profiles.Sort((x, y) => string.Compare(x.name, y.name));
        }

        public static void sort_commands_by_name_asc(int profile_ind)
        {
            profiles[profile_ind].custom_commands.Sort((x, y) => string.Compare(x.name, y.name));
        }

        public static void sort_groups_by_name_asc()
        {
            groups.Sort((x, y) => string.Compare(x.name, y.name));
        }

        public static VirtualKeyCode get_virtual_key_code_by_key_name(string key_name)
        {
            foreach(VirtualKey vk in keys)
            {
                if (vk.name == key_name)
                    return vk.vkc;
            }

            return VirtualKeyCode.VK_V;
        }

        public static string get_key_name_by_virtual_key_code(VirtualKeyCode vkc)
        {
            foreach (VirtualKey vk in keys)
            {
                if (vk.vkc == vkc)
                    return vk.name;
            }

            return "";
        }

        public static int get_profile_ind_by_name(string name)
        {
            for(int i=0; i< profiles.Count; i++)
            {
                if (profiles[i].name == name)
                    return i;
            }

            return -1;
        }

        public static int get_command_ind_by_name(int profile_ind, string name)
        {
            for (int i = 0; i < profiles[profile_ind].custom_commands.Count; i++)
            {
                if (profiles[profile_ind].custom_commands[i].name == name)
                    return i;
            }

            return -1;
        }

        public static string get_command_group_by_name(int profile_ind, string name)
        {
            for (int i = 0; i < profiles[profile_ind].custom_commands.Count; i++)
            {
                if (profiles[profile_ind].custom_commands[i].name == name)
                    return profiles[profile_ind].custom_commands[i].group;
            }

            return general_group_name;
        }

        public static int get_nr_of_commands_in_profiles_for_specific_apps()
        {
            int n = 0;

            foreach (Profile p in profiles)
            {
                if(p.name != Middle_Man.any_program_name)
                    n += p.custom_commands.Count;
            }

            return n;
        }

        public static int get_nr_of_commands_in_all_profiles()
        {
            int n = 0;

            foreach(Profile p in profiles)
            {
                n += p.custom_commands.Count;
            }

            return n;
        }

        public static bool contains_illegal_characers_or_names(string str)
        {
            str = str.ToUpper();

            if (str.Contains("\\")
                || str.Contains("/")
                || str.Contains(":")
                || str.Contains("?")
                || str.Contains("*")
                || str.Contains("\"")
                || str.Contains("<")
                || str.Contains(">")
                || str.Contains("|")
                || str == "CON"
                || str == "PRN"
                || str == "AUX"
                || str == "NUL"
                || str == "COM1"
                || str == "COM2"
                || str == "COM3"
                || str == "COM4"
                || str == "COM5"
                || str == "COM6"
                || str == "COM7"
                || str == "COM8"
                || str == "COM9"
                || str == "LPT1"
                || str == "LPT2"
                || str == "LPT3"
                || str == "LPT4"
                || str == "LPT5"
                || str == "LPT6"
                || str == "LPT7"
                || str == "LPT8"
                || str == "LPT9")
            {
                return true;
            }
            else 
                return false;
        }        
    }
}
