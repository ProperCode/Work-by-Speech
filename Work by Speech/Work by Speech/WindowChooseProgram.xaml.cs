using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace Speech
{
    /// <summary>
    /// Interaction logic for Full_version_activation.xaml
    /// </summary>
    public partial class WindowChooseProgram : Window
    {
        class Program
        {
            public BitmapImage process_image { get; set; }
            public string process_name { get; set; }
            public string window_title { get; set; }
        }

        List<Program> programs = new List<Program>();

        public WindowChooseProgram()
        {
            try
            {
                InitializeComponent();

                LVprograms.SelectionMode = System.Windows.Controls.SelectionMode.Single;

                load_open_programs();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WCP001", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void load_open_programs()
        {
            try
            {
                programs = new List<Program>();

                LVprograms.ItemsSource = programs;

                Icon ico;
                Bitmap bi;

                bi = DrawFilledRectangle(32, 32);

                programs.Add(new Program()
                {
                    process_image = BitmapToImageSource(bi),
                    process_name = Middle_Man.any_program_name,
                    window_title = ""
                }); ;

                Process[] arr = Process.GetProcesses();

                for (int i = 0; i < arr.Length; i++)
                {
                    if (String.IsNullOrEmpty(arr[i].MainWindowTitle) == false)
                    {
                        ico = System.Drawing.Icon.ExtractAssociatedIcon(arr[i].MainModule.FileName);
                        bi = ico.ToBitmap();
                        bi = Transparent2Color(bi, Color.White);

                        programs.Add(new Program()
                        {
                            process_image = BitmapToImageSource(bi),
                            process_name = arr[i].ProcessName,
                            window_title = arr[i].MainWindowTitle
                        });
                    }
                }

                CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(LVprograms.ItemsSource);
                cv.SortDescriptions.Add(new System.ComponentModel.SortDescription("process_name",
                    System.ComponentModel.ListSortDirection.Ascending));

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WCP002", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                LVprograms.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WCP003", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Brefresh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                load_open_programs();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WCP004", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Bok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (System.Windows.Window window in Application.Current.Windows)
                {
                    if (window.GetType() == typeof(WindowAddEditProfile))
                    {
                        if (LVprograms.SelectedIndex != -1)
                        {
                            WindowAddEditProfile w = (WindowAddEditProfile)window;
                            w.TBprogram.Text = ((Program)LVprograms.SelectedItem).process_name;
                        }
                    }
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WCP005", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "Error WCP006", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LVprograms_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    Bok_Click(null, null);
                }
                else if (e.Key == Key.Escape)
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error WCP007", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }

        Bitmap Transparent2Color(Bitmap bmp1, Color target)
        {
            Bitmap bmp2 = new Bitmap(bmp1.Width, bmp1.Height);
            Rectangle rect = new Rectangle(System.Drawing.Point.Empty, bmp1.Size);
            using (Graphics G = Graphics.FromImage(bmp2))
            {
                G.Clear(target);
                G.DrawImageUnscaledAndClipped(bmp1, rect);
            }
            return bmp2;
        }

        Bitmap DrawFilledRectangle(int x, int y)
        {
            Bitmap bmp = new Bitmap(x, y);
            using (Graphics graph = Graphics.FromImage(bmp))
            {
                Rectangle ImageSize = new Rectangle(0, 0, x, y);
                graph.FillRectangle(Brushes.White, ImageSize);
            }
            return bmp;
        }
    }
}