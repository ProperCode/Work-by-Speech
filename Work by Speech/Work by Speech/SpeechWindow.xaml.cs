using System;
using System.ComponentModel;
using System.IO;
using System.Windows;

namespace Speech
{
    /// <summary>
    /// Interaction logic for SpeechWindow.xaml
    /// </summary>
    public partial class SpeechWindow : Window
    {
        public int mode = 0;
        public bool change_mode = false;

        const string filename_coords = "coords.txt"; //speech recognition window last location
        string app_folder_path = System.IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location);

        public SpeechWindow()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error SW001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void Bmode_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                change_mode = true;

                if (Bmode.Content.ToString() == "OFF")
                {
                    mode = 1;
                }
                else if (Bmode.Content.ToString() == "Command")
                {
                    mode = 2;
                }
                else mode = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error SW002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
}

        private void Window_LocationChanged(object sender, EventArgs e)
        {
            if (this.IsInitialized)
            {
                FileStream fs = null;
                StreamWriter sw = null;
                string file_path = System.IO.Path.Combine(app_folder_path, filename_coords);

                try
                {
                    fs = new FileStream(file_path, FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs);

                    sw.WriteLine(((int)this.Left).ToString());
                    sw.WriteLine(((int)this.Top).ToString());

                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error SW003", MessageBoxButton.OK, MessageBoxImage.Error);

                    try
                    {
                        if (sw != null)
                            sw.Close();
                        if (fs != null)
                            fs.Close();
                    }
                    catch (Exception ex2) { }
                }
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            try
            {
                base.OnClosing(e);
                e.Cancel = true;
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error SW004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}