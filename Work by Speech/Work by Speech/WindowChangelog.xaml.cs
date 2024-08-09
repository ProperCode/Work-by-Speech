using System;
using System.Windows;

namespace Speech
{
    /// <summary>
    /// Interaction logic for EULA.xaml
    /// </summary>
    public partial class WindowChangelog : Window
    {
        public WindowChangelog()
        {
            try
            {
                InitializeComponent();

                TB.IsReadOnly = true;

                TB.Text = "All notable changes to Work by Speech will be documented here."
                + "\n\n[2.2] - August 10, 2024:"
                + "\n- Fixed access denied error which was ocurring in non-US Windows 11 installations."
                + "\n\n[2.1] - January 25, 2024:"
                + "\n- Fixed smart mousegrid and mode switching."
                + "\n\n[2.0] - December 27, 2023:"
                + "\n- Work by Speech is from now on an open source application."
                + "\n\n[1.9] - October 11, 2023:"
                + "\n- Fixed rare freezing."
                + "\n\n[1.8] - September 15, 2023:"
                + "\n- Fixed a minor bug."
                + "\n- Improved macro recording."                
                + "\n- Other minor improvements."
                + "\n\n[1.7] - September 11, 2023:"
                + "\n- Added macro recording system."
                + "\n- Added built-in commands: start recording, stop recording."
                + "\n- Added speech synthesis volume changing."
                + "\n- Fixed a rare bug and minor bugs."
                + "\n\n[1.6] - June 30, 2023:"
                + "\n- Added built-in commands toggling system."
                + "\n- Added built-in command: web address."
                + "\n- Changing mousegrid type or desired figures number no longer deletes " +
                "Smart mousegrid data."
                + "\n- Fixed rare bugs."
                + "\n- Improved settings saving."
                + "\n\n[1.5] - May 31, 2023:"
                + "\n- Fixed custom command action \"Release all buttons and keys\" not releasing" +
                " mouse buttons."
                + "\n- Fixed minor bugs."
                + "\n- Improved UI."
                + "\n\n[1.4] - May 24, 2023:"
                + "\n- Added built-in command: get position."
                + "\n- Added custom commands management system."
                + "\n- Changed how key combinations work for built-in commands."
                + "\n- Improved mousegrid and UI."
                + "\n- Removed less useful character inserting commands."
                + "\n\n[1.3] - March 6, 2023:"
                + "\n- Added new commands: triple, triple click."
                + "\n- Added speech synthesis voice selection."
                + "\n- Changed alt, control, shift, windows maximum executions to 20."
                + "\n- Changed left, right maximum executions to 90."
                + "\n- Fixed repeated key combinations execution."
                + "\n- Fixed a rare settings saving bug (infinite waiting for saving)."
                + "\n- Improved settings saving."
                + "\n- Print screen can now be used in key combinations."
                + "\n\n[1.2] - February 24, 2023:"
                + "\n- Fixed a rare mode changing bug."
                + "\n- Improved performance, settings saving and error handling."
                + "\n- Moved semicolon, quote / apostrophe, backslash, open bracket and close bracket"
                + " from keys pressing commands to character inserting commands to allow inserting these" +
                " special characters on all keyboard layouts."
                + "\n- Removed automatic updating."
                + "\n\n[1.1] - February 18, 2023:"
                + "\n- Improved automatic updating.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WC001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}