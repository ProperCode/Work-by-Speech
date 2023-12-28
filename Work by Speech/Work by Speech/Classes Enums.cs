using System.Collections.Generic;
using System.Windows;

namespace Speech
{
    public class Grid_element
	{
		public string symbol;
		public string word;
		public int word_length;
		public uint count = 0;

		public Grid_element(string w, string s)
		{
			symbol = s;
			word = w;
			word_length = w.Length;
		}
	}

	public partial class MainWindow : Window
    {
		public class Process_grid
		{
			public string process_name;
			public List<Grid_element> elements = new List<Grid_element>();
			public int count;

			public Process_grid(string Process_name)
			{
				process_name = Process_name;
			}
		}

		public class Grid_Symbol
		{
			public string symbol;
			public string word;

			public Grid_Symbol(string w, string s = null)
			{
				if (s == null) s = w[0].ToString();
				symbol = s;
				word = w;
			}
		}
				
		public class Installed_App
		{
			public List<string> names;
			public int name_length;
			public string path;
		}

		enum mode
		{
			off,
			command,
			dictation,
			grid //part of command mode
		}

        enum grammar_type
        {
            grammar_off_mode,
            grammar_dictation_commands,
            grammar_dictation,
            grammar_mousegrid,
            grammar_builtin_commands,
            grammar_custom_commands_any, //any program
            grammar_custom_commands_foreground, //foreground program
            grammar_apps_switching,
            grammar_apps_opening
        }

        public enum bic_type //bic = built-in command
        {
			turn_on, //speech recognition
			turn_off, //speech recognition
			switch_to_command, //mode
            switch_to_dictation, //mode
            show_speech_recognition,
            hide_speech_recognition,
            start_better_dictation_listening,
            toggle_better_dictation,
            close_that,
            minimize_that,
            maximize_that,
            restore_that,
            switch_to_app,
            open_app,
			get_position,
			start_recording,
            stop_recording,
            key_combination,
            key_pressing,
			character_ins, //character inserting
            cancel,
            release_buttons,
            move,
            left,
            right,
            double2,
            triple,
            drag,
            drop,
            click,
            right_click,
            double_click,
            triple_click,
            hold,
            hold_right,
            move_up,
            move_down,
            move_left,
            move_right,
            scroll_up,
            scroll_down,
            scroll_left,
            scroll_right,
			move_top_edge,
            move_bottom_edge,
            move_left_edge,
            move_right_edge,
            move_screen_center
        }

        public enum GridType
        {
            hexagonal,
            square,
            square_horizontal_precision,
            square_vertical_precision,
            square_combined_precision
        }

        void create_optimized_grid_alphabet()
		{
			grid_alphabet = new List<Grid_Symbol>();

            grid_alphabet.Add(new Grid_Symbol("alfa", "a"));
            grid_alphabet.Add(new Grid_Symbol("bravo", "b"));
            grid_alphabet.Add(new Grid_Symbol("charlie", "c"));
            grid_alphabet.Add(new Grid_Symbol("delta", "d"));
            grid_alphabet.Add(new Grid_Symbol("echo", "e"));
            grid_alphabet.Add(new Grid_Symbol("foxtrot", "f"));
            grid_alphabet.Add(new Grid_Symbol("golf", "g"));
            grid_alphabet.Add(new Grid_Symbol("hotel", "h"));
            grid_alphabet.Add(new Grid_Symbol("india", "i"));
            grid_alphabet.Add(new Grid_Symbol("juliett", "j"));
            grid_alphabet.Add(new Grid_Symbol("kilo", "k"));
            grid_alphabet.Add(new Grid_Symbol("lima", "l"));
            grid_alphabet.Add(new Grid_Symbol("november", "n"));
            grid_alphabet.Add(new Grid_Symbol("oscar", "o"));
            grid_alphabet.Add(new Grid_Symbol("papa", "p"));
            grid_alphabet.Add(new Grid_Symbol("quebec", "q"));
            grid_alphabet.Add(new Grid_Symbol("romeo", "r"));
            grid_alphabet.Add(new Grid_Symbol("sierra", "s"));
            grid_alphabet.Add(new Grid_Symbol("tango", "t"));
            grid_alphabet.Add(new Grid_Symbol("uniform", "u"));
            grid_alphabet.Add(new Grid_Symbol("victor", "v"));
            grid_alphabet.Add(new Grid_Symbol("xray", "x"));
            grid_alphabet.Add(new Grid_Symbol("yankee", "y"));
            grid_alphabet.Add(new Grid_Symbol("zulu", "z"));
            grid_alphabet.Add(new Grid_Symbol("one", "1"));
            grid_alphabet.Add(new Grid_Symbol("two", "2"));
            grid_alphabet.Add(new Grid_Symbol("three", "3"));
            grid_alphabet.Add(new Grid_Symbol("four", "4"));
            grid_alphabet.Add(new Grid_Symbol("five", "5"));
            grid_alphabet.Add(new Grid_Symbol("six", "6"));
            grid_alphabet.Add(new Grid_Symbol("seven", "7"));
            grid_alphabet.Add(new Grid_Symbol("eight", "8"));
            grid_alphabet.Add(new Grid_Symbol("nine", "9"));
            grid_alphabet.Add(new Grid_Symbol("slash", "/"));
			grid_alphabet.Add(new Grid_Symbol("minus", "-"));
			grid_alphabet.Add(new Grid_Symbol("question", "?"));
            grid_alphabet.Add(new Grid_Symbol("pound", "£"));
            grid_alphabet.Add(new Grid_Symbol("dollar", "$"));     
			grid_alphabet.Add(new Grid_Symbol("euro", "€"));
			grid_alphabet.Add(new Grid_Symbol("yen", "¥"));
			grid_alphabet.Add(new Grid_Symbol("asterisk", "*"));
			grid_alphabet.Add(new Grid_Symbol("ampersand", "&"));
			grid_alphabet.Add(new Grid_Symbol("bracket", "]"));
			grid_alphabet.Add(new Grid_Symbol("brace", "{"));
			grid_alphabet.Add(new Grid_Symbol("section", "§"));
			grid_alphabet.Add(new Grid_Symbol("function", "ƒ"));
			grid_alphabet.Add(new Grid_Symbol("degree", "°"));
			grid_alphabet.Add(new Grid_Symbol("quote", "\""));
			grid_alphabet.Add(new Grid_Symbol("colon", ":"));
			grid_alphabet.Add(new Grid_Symbol("sigma", "∑"));
			
			//grid_alphabet.Add(new Grid_Symbol("comma", ",")); //a bit hard to see
			//grid_alphabet.Add(new Grid_Symbol("paragraph", "¶")); //too long code word

			//a bit too wide:
			//grid_alphabet.Add(new Grid_Symbol("whiskey", "w"));
			//grid_alphabet.Add(new Grid_Symbol("plus", "+"));
			//grid_alphabet.Add(new Grid_Symbol("equal", "="));
			//grid_alphabet.Add(new Grid_Symbol("hash", "#"));
			//grid_alphabet.Add(new Grid_Symbol("caret", "^"));

			//grid_alphabet.Add(new Grid_Symbol("backslash", "\\"));
			//grid_alphabet.Add(new Grid_Symbol("open paren", "("));
			//grid_alphabet.Add(new Grid_Symbol("close paren", ")"));

			//too wide:
			//grid_alphabet.Add(new Grid_Symbol("less than", "<"));
			//grid_alphabet.Add(new Grid_Symbol("tilde", "~"));
			//grid_alphabet.Add(new Grid_Symbol("mike", "m"));
		}

		void create_normal_grid_alphabet()
		{
			grid_alphabet = new List<Grid_Symbol>();

			grid_alphabet.Add(new Grid_Symbol("alfa", "a"));
			grid_alphabet.Add(new Grid_Symbol("bravo", "b"));
			grid_alphabet.Add(new Grid_Symbol("charlie", "c"));
			grid_alphabet.Add(new Grid_Symbol("delta", "d"));
			grid_alphabet.Add(new Grid_Symbol("echo", "e"));
			grid_alphabet.Add(new Grid_Symbol("foxtrot", "f"));
			grid_alphabet.Add(new Grid_Symbol("golf", "g"));
			grid_alphabet.Add(new Grid_Symbol("hotel", "h"));
			grid_alphabet.Add(new Grid_Symbol("india", "i"));
			grid_alphabet.Add(new Grid_Symbol("juliett", "j"));
			grid_alphabet.Add(new Grid_Symbol("kilo", "k"));
			grid_alphabet.Add(new Grid_Symbol("lima", "l"));
			grid_alphabet.Add(new Grid_Symbol("mike", "m"));
			grid_alphabet.Add(new Grid_Symbol("november", "n"));
			grid_alphabet.Add(new Grid_Symbol("oscar", "o"));
			grid_alphabet.Add(new Grid_Symbol("papa", "p"));
			grid_alphabet.Add(new Grid_Symbol("quebec", "q"));
			grid_alphabet.Add(new Grid_Symbol("romeo", "r"));
			grid_alphabet.Add(new Grid_Symbol("sierra", "s"));
			grid_alphabet.Add(new Grid_Symbol("tango", "t"));
			grid_alphabet.Add(new Grid_Symbol("uniform", "u"));
			grid_alphabet.Add(new Grid_Symbol("victor", "v"));
			grid_alphabet.Add(new Grid_Symbol("whiskey", "w"));
			grid_alphabet.Add(new Grid_Symbol("xray", "x"));
			grid_alphabet.Add(new Grid_Symbol("yankee", "y"));
			grid_alphabet.Add(new Grid_Symbol("zulu", "z"));
			grid_alphabet.Add(new Grid_Symbol("one", "1"));
			grid_alphabet.Add(new Grid_Symbol("two", "2"));
			grid_alphabet.Add(new Grid_Symbol("three", "3"));
			grid_alphabet.Add(new Grid_Symbol("four", "4"));
			grid_alphabet.Add(new Grid_Symbol("five", "5"));
			grid_alphabet.Add(new Grid_Symbol("six", "6"));
			grid_alphabet.Add(new Grid_Symbol("seven", "7"));
			grid_alphabet.Add(new Grid_Symbol("eight", "8"));
			grid_alphabet.Add(new Grid_Symbol("nine", "9"));
			grid_alphabet.Add(new Grid_Symbol("zero", "0"));
		}

		void create_wide_grid_alphabet()
		{
			grid_alphabet = new List<Grid_Symbol>();

			grid_alphabet.Add(new Grid_Symbol("mike", "m"));
			grid_alphabet.Add(new Grid_Symbol("whiskey", "w"));
			grid_alphabet.Add(new Grid_Symbol("Tilde", "~"));
			grid_alphabet.Add(new Grid_Symbol("At", "@"));
			grid_alphabet.Add(new Grid_Symbol("Hash", "#"));
			grid_alphabet.Add(new Grid_Symbol("Percent", "%"));
			grid_alphabet.Add(new Grid_Symbol("Ampersand", "&"));
			grid_alphabet.Add(new Grid_Symbol("Section", "§"));
			grid_alphabet.Add(new Grid_Symbol("Paragraph", "¶"));
			grid_alphabet.Add(new Grid_Symbol("Function", "ƒ"));
			grid_alphabet.Add(new Grid_Symbol("Micro", "µ"));
		}

		void create_full_grid_alphabet()
        {
			grid_alphabet = new List<Grid_Symbol>();

			grid_alphabet.Add(new Grid_Symbol("alfa", "a"));
			grid_alphabet.Add(new Grid_Symbol("bravo", "b"));
			grid_alphabet.Add(new Grid_Symbol("charlie", "c"));
			grid_alphabet.Add(new Grid_Symbol("delta", "d"));
			grid_alphabet.Add(new Grid_Symbol("echo", "e"));
			grid_alphabet.Add(new Grid_Symbol("foxtrot", "f"));
			grid_alphabet.Add(new Grid_Symbol("golf", "g"));
			grid_alphabet.Add(new Grid_Symbol("hotel", "h"));
			grid_alphabet.Add(new Grid_Symbol("india", "i"));
			grid_alphabet.Add(new Grid_Symbol("juliett", "j"));
			grid_alphabet.Add(new Grid_Symbol("kilo", "k"));
			grid_alphabet.Add(new Grid_Symbol("lima", "l"));
			grid_alphabet.Add(new Grid_Symbol("mike", "m"));
			grid_alphabet.Add(new Grid_Symbol("november", "n"));
			grid_alphabet.Add(new Grid_Symbol("oscar", "o"));
			grid_alphabet.Add(new Grid_Symbol("papa", "p"));
			grid_alphabet.Add(new Grid_Symbol("quebec", "q"));
			grid_alphabet.Add(new Grid_Symbol("romeo", "r"));
			grid_alphabet.Add(new Grid_Symbol("sierra", "s"));
			grid_alphabet.Add(new Grid_Symbol("tango", "t"));
			grid_alphabet.Add(new Grid_Symbol("uniform", "u"));
			grid_alphabet.Add(new Grid_Symbol("victor", "v"));
			grid_alphabet.Add(new Grid_Symbol("whiskey", "w"));
			grid_alphabet.Add(new Grid_Symbol("xray", "x"));
			grid_alphabet.Add(new Grid_Symbol("yankee", "y"));
			grid_alphabet.Add(new Grid_Symbol("zulu", "z"));
			grid_alphabet.Add(new Grid_Symbol("one", "1"));
			grid_alphabet.Add(new Grid_Symbol("two", "2"));
			grid_alphabet.Add(new Grid_Symbol("three", "3"));
			grid_alphabet.Add(new Grid_Symbol("four", "4"));
			grid_alphabet.Add(new Grid_Symbol("five", "5"));
			grid_alphabet.Add(new Grid_Symbol("six", "6"));
			grid_alphabet.Add(new Grid_Symbol("seven", "7"));
			grid_alphabet.Add(new Grid_Symbol("eight", "8"));
			grid_alphabet.Add(new Grid_Symbol("nine", "9"));
			grid_alphabet.Add(new Grid_Symbol("zero", "0"));
			grid_alphabet.Add(new Grid_Symbol("Comma", ","));
			grid_alphabet.Add(new Grid_Symbol("Semicolon", ";"));
			grid_alphabet.Add(new Grid_Symbol("Dot", "."));
			grid_alphabet.Add(new Grid_Symbol("Quote", "'"));
			grid_alphabet.Add(new Grid_Symbol("Slash", "/"));
			grid_alphabet.Add(new Grid_Symbol("Backslash", "\\"));
			grid_alphabet.Add(new Grid_Symbol("Minus", "-"));
			grid_alphabet.Add(new Grid_Symbol("Open bracket", "["));
			grid_alphabet.Add(new Grid_Symbol("Close bracket", "]"));
			grid_alphabet.Add(new Grid_Symbol("Asterisk", "*"));
			grid_alphabet.Add(new Grid_Symbol("Plus", "+"));
			grid_alphabet.Add(new Grid_Symbol("Equal", "="));
			grid_alphabet.Add(new Grid_Symbol("Colon", ":"));
			grid_alphabet.Add(new Grid_Symbol("Double quote", "\""));
			grid_alphabet.Add(new Grid_Symbol("Greater than", ">"));
			grid_alphabet.Add(new Grid_Symbol("Less than", "<"));
			grid_alphabet.Add(new Grid_Symbol("Tilde", "~"));
			grid_alphabet.Add(new Grid_Symbol("At", "@"));
			grid_alphabet.Add(new Grid_Symbol("Exclamation", "!"));
			grid_alphabet.Add(new Grid_Symbol("Question", "?"));
			grid_alphabet.Add(new Grid_Symbol("Hash", "#"));;
			grid_alphabet.Add(new Grid_Symbol("Pound", "£"));
			grid_alphabet.Add(new Grid_Symbol("Dollar", "$"));
			grid_alphabet.Add(new Grid_Symbol("Percent", "%"));
			grid_alphabet.Add(new Grid_Symbol("Caret", "^"));
			grid_alphabet.Add(new Grid_Symbol("Open paren", "("));
			grid_alphabet.Add(new Grid_Symbol("Close paren", ")"));
			grid_alphabet.Add(new Grid_Symbol("Underscore", "_"));
			grid_alphabet.Add(new Grid_Symbol("Open brace", "{"));
			grid_alphabet.Add(new Grid_Symbol("Close brace", "}"));
			grid_alphabet.Add(new Grid_Symbol("Vertical bar", "|"));
			grid_alphabet.Add(new Grid_Symbol("Trademark", "™"));
			grid_alphabet.Add(new Grid_Symbol("Three-quarter", "¾"));
			grid_alphabet.Add(new Grid_Symbol("One-quarter", "¼"));
			grid_alphabet.Add(new Grid_Symbol("One-half", "½"));
			grid_alphabet.Add(new Grid_Symbol("Ampersand", "&"));
			grid_alphabet.Add(new Grid_Symbol("Back quote", "`"));
			grid_alphabet.Add(new Grid_Symbol("Plus or minus", "±"));
			grid_alphabet.Add(new Grid_Symbol("Open angle quote", "«"));
			grid_alphabet.Add(new Grid_Symbol("Close angle quote", "»"));
			grid_alphabet.Add(new Grid_Symbol("Division", "÷"));
			grid_alphabet.Add(new Grid_Symbol("Cent", "¢"));
			grid_alphabet.Add(new Grid_Symbol("Yen", "¥"));
			grid_alphabet.Add(new Grid_Symbol("Section", "§"));
			grid_alphabet.Add(new Grid_Symbol("Copyright", "©"));
			grid_alphabet.Add(new Grid_Symbol("Registered", "®"));
			grid_alphabet.Add(new Grid_Symbol("Degree", "°"));
			grid_alphabet.Add(new Grid_Symbol("Paragraph", "¶"));
			grid_alphabet.Add(new Grid_Symbol("Function", "ƒ"));
			grid_alphabet.Add(new Grid_Symbol("Micro", "µ"));
		}		
	}
}