using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Speech
{
    /// <summary>
    /// Interaction logic for MouseGrid.xaml
    /// </summary>
    public partial class MouseGrid : Window
    {
        MainWindow.GridType grid_type;
        int rows_nr;
        int cols_nr;
        double figure_width;
        double figure_height;
        FontFamily font_family;
        int font_size;
        Color color1;
        Color color2;
        public List<Grid_element> elements;

        List<TextBlock> list_tb = new List<TextBlock>();
        Canvas canvas;

        public MouseGrid(double grid_width, double grid_height, int grid_lines, MainWindow.GridType Grid_type,
            FontFamily Font_family, int Font_size, Color Color1, Color Color2, int Rows_nr, int Cols_nr,
            double Figure_width, double Figure_height, List<Grid_element> Elements)
        {
            InitializeComponent();

            int scaling = GetWindowsScaling();

            grid_type = Grid_type;
            rows_nr = Rows_nr;
            cols_nr = Cols_nr;
            font_family = Font_family;
            color1 = Color1;
            color2 = Color2;
            elements = Elements;

            figure_width = Figure_width;
            figure_height = Figure_height;
            //font_size = Font_size;
            font_size = (int)(Font_size * 100 / scaling);

            this.Topmost = true;
            this.ShowInTaskbar = false;

            int ind = 0;
            Random random = new Random();
            Size size;
            
            if (grid_type==MainWindow.GridType.hexagonal)
            {
                canvas = new Canvas();
                canvas.Width = this.Width = grid_width;
                canvas.Height = this.Height = grid_height;
                canvas.HorizontalAlignment = HorizontalAlignment.Center;
                canvas.VerticalAlignment = VerticalAlignment.Center;
                canvas.AllowDrop = false;
                canvas.SnapsToDevicePixels = true;

                double even_row_margin_left = Math.Sqrt(Math.Pow(figure_width, 2)
                    - Math.Pow(0.5 * figure_width, 2)) * 0.5;
                even_row_margin_left = figure_width * 0.5;
                double even_row_margin_top = figure_height * 0.5;

                int rows_nr2 = rows_nr * 2 - 1;

                for (int i = 0; i < rows_nr2; i++)
                {
                    for (int j = 0; j < cols_nr; j++)
                    {
                        if (i % 2 == 0 || j < cols_nr - 1)
                        {
                            list_tb.Add(new TextBlock());
                            list_tb[ind].Text = elements[ind].symbol;
                            //list_tb[ind].Text = random.Next(0, 100).ToString();
                            //list_tb[ind].Text = "a5";
                            list_tb[ind].Foreground = new SolidColorBrush(color2);
                            list_tb[ind].Background = new SolidColorBrush(color1);
                            list_tb[ind].FontSize = font_size;
                            list_tb[ind].FontFamily = font_family;
                            //list_tb[ind].FontWeight = FontWeights.Bold;
                            list_tb[ind].TextAlignment = TextAlignment.Center;
                            list_tb[ind].VerticalAlignment = VerticalAlignment.Center;
                            size = MeasureString(list_tb[ind].Text, list_tb[ind]);
                            list_tb[ind].Width = size.Width + 2;
                            list_tb[ind].Height = size.Height + 2;

                            canvas.Children.Add(list_tb[ind]);

                            if (i % 2 == 0)
                            {
                                Canvas.SetLeft(list_tb[ind], (j + 0.5) * figure_width
                                    - list_tb[ind].Width / 2);
                                Canvas.SetTop(list_tb[ind], ((int)(i / 2) + 0.5) * figure_height
                                    - list_tb[ind].Height / 2);
                            }
                            else
                            {
                                Canvas.SetLeft(list_tb[ind], even_row_margin_left + (j + 0.5) * figure_width
                                - list_tb[ind].Width / 2);
                                Canvas.SetTop(list_tb[ind], even_row_margin_top + ((int)(i / 2) + 0.5)
                                    * figure_height - list_tb[ind].Height / 2);
                            }
                            
                            ind++;
                        }
                    }
                }

                this.Content = canvas;

                //unnecessary:
                //if (scaling <= 100)
                //{
                //    this.Content = canvas;
                //}
                ////descaling
                //else
                //{
                //    // Set the center point of the ##.
                //    canvas.RenderTransformOrigin = new Point(0, 0);

                //    // Create a transform to scale the size of the button.
                //    ScaleTransform myScaleTransform = new ScaleTransform();

                //    myScaleTransform.ScaleX = myScaleTransform.ScaleY = 100 / (double)scaling;

                //    // Create a TransformGroup to contain the transforms
                //    // and add the transforms to it.
                //    TransformGroup myTransformGroup = new TransformGroup();
                //    myTransformGroup.Children.Add(myScaleTransform);

                //    // Associate the transforms to the button.
                //    canvas.RenderTransform = myTransformGroup;

                //    // Create a StackPanel which will contain the Button.
                //    //StackPanel myStackPanel = new StackPanel();
                //    //myStackPanel.Children.Add(canvas);
                //    this.Content = canvas;
                //}
            }
            else
            {
                // Create the Grid
                Grid g = new Grid();
                this.Width = g.Width = grid_width;
                this.Height = g.Height = grid_height;
                g.HorizontalAlignment = HorizontalAlignment.Center;
                g.VerticalAlignment = VerticalAlignment.Center;
                if (grid_lines == 1)
                    g.ShowGridLines = true;
                else
                    g.ShowGridLines = false;
                g.AllowDrop = false;
                g.SnapsToDevicePixels = true;

                List<ColumnDefinition> list_cols = new List<ColumnDefinition>();
                List<RowDefinition> list_rows = new List<RowDefinition>();
                List<Rectangle> list_r = new List<Rectangle>();
                List<Rectangle> list_r2 = new List<Rectangle>();

                for (int i = 0; i < cols_nr; i++)
                {
                    list_cols.Add(new ColumnDefinition());
                    g.ColumnDefinitions.Add(list_cols[i]);
                }
                for (int i = 0; i < rows_nr; i++)
                {
                    list_rows.Add(new RowDefinition());
                    g.RowDefinitions.Add(list_rows[i]);
                }

                for (int i = 0; i < rows_nr; i++)
                {
                    for (int j = 0; j < cols_nr; j++)
                    { 
                        list_tb.Add(new TextBlock());
                        list_tb[ind].Text = elements[ind].symbol;
                        //list_tb[ind].Text = random.Next(0, 100).ToString();// font_size.ToString();// "a5";
                        list_tb[ind].Foreground = new SolidColorBrush(color2);
                        list_tb[ind].Background = new SolidColorBrush(color1);
                        list_tb[ind].FontSize = font_size;
                        list_tb[ind].FontFamily = font_family;
                        //list_tb[ind].FontWeight = FontWeights.Bold;
                        
                        size = MeasureString(list_tb[ind].Text, list_tb[ind]);
                        list_tb[ind].Width = size.Width + 2;
                        list_tb[ind].Height = size.Height + 2;

                        list_tb[ind].TextAlignment = TextAlignment.Center;
                        list_tb[ind].VerticalAlignment = VerticalAlignment.Center;
                        list_tb[ind].HorizontalAlignment = HorizontalAlignment.Center;

                        if (grid_type == MainWindow.GridType.square_horizontal_precision
                            || grid_type == MainWindow.GridType.square_combined_precision)
                        {
                            if (j % 3 == 0)
                                list_tb[ind].VerticalAlignment = VerticalAlignment.Top;
                            else if (j % 3 == 1)
                                list_tb[ind].VerticalAlignment = VerticalAlignment.Center;
                            else
                                list_tb[ind].VerticalAlignment = VerticalAlignment.Bottom;
                        }
                        if (grid_type == MainWindow.GridType.square_vertical_precision
                            || grid_type == MainWindow.GridType.square_combined_precision)
                        {
                            if (i % 3 == 0)
                                list_tb[ind].HorizontalAlignment = HorizontalAlignment.Left;
                            else if (i % 3 == 1)
                                list_tb[ind].HorizontalAlignment = HorizontalAlignment.Center;
                            else
                                list_tb[ind].HorizontalAlignment = HorizontalAlignment.Right;
                        }

                        Grid.SetRow(list_tb[ind], i);
                        Grid.SetColumn(list_tb[ind], j);

                        g.Children.Add(list_tb[ind]);

                        if (grid_lines == 2)
                        {
                            list_r.Add(new Rectangle());
                            list_r[ind].Fill = Brushes.Transparent;
                            list_r[ind].Stroke = new SolidColorBrush(color1);
                            list_r[ind].StrokeThickness = 1;
                            list_r[ind].Width = figure_width;
                            list_r[ind].Height = figure_height;

                            Grid.SetRow(list_r[ind], i);
                            Grid.SetColumn(list_r[ind], j);

                            g.Children.Add(list_r[ind]);

                            list_r2.Add(new Rectangle());
                            list_r2[ind].Fill = Brushes.Transparent;
                            list_r2[ind].Stroke = new SolidColorBrush(color2);
                            list_r2[ind].StrokeThickness = 1;
                            list_r2[ind].Width = figure_width;
                            list_r2[ind].Height = figure_height;

                            Grid.SetRow(list_r2[ind], i);
                            Grid.SetColumn(list_r2[ind], j);

                            g.Children.Add(list_r2[ind]);
                        }

                        ind++;
                    }
                }

                // Add the Grid as the Content of the Parent Window Object
                this.Content = g;

                //unnecessary:
                //if (scaling <= 100)
                //{
                //    // Add the Grid as the Content of the Parent Window Object
                //    this.Content = g;
                //}
                ////descaling
                //else
                //{
                //    // Set the center point of the transforms.
                //    g.RenderTransformOrigin = new Point(0, 0);

                //    // Create a transform to scale the size of the button.
                //    ScaleTransform myScaleTransform = new ScaleTransform();

                //    myScaleTransform.ScaleX = myScaleTransform.ScaleY = 100 / (double)scaling;

                //    // Create a TransformGroup to contain the transforms
                //    // and add the transforms to it.
                //    TransformGroup myTransformGroup = new TransformGroup();
                //    myTransformGroup.Children.Add(myScaleTransform);

                //    // Associate the transforms to the button.
                //    g.RenderTransform = myTransformGroup;

                //    // Create a StackPanel which will contain the Button.
                //    //StackPanel myStackPanel = new StackPanel();
                //    //myStackPanel.Children.Add(canvas);
                //    this.Content = g;
                //}
            }

            CenterWindowOnScreen();
        }

        public void regenerate_grid_symbols()
        {
            int ind = 0;
            Random random = new Random();
            Size size;

            if (grid_type == MainWindow.GridType.hexagonal)
            {
                double even_row_margin_left = Math.Sqrt(Math.Pow(figure_width, 2)
                    - Math.Pow(0.5 * figure_width, 2)) * 0.5;
                even_row_margin_left = figure_width * 0.5;
                double even_row_margin_top = figure_height * 0.5;

                int rows_nr2 = rows_nr * 2 - 1;

                for (int i = 0; i < rows_nr2; i++)
                {
                    for (int j = 0; j < cols_nr; j++)
                    {
                        if (i % 2 == 0 || j < cols_nr - 1)
                        {
                            if (list_tb[ind].Text != elements[ind].symbol)
                            {
                                list_tb[ind].Text = elements[ind].symbol;
                                //list_tb[ind].Text = random.Next(0, 100).ToString();// font_size.ToString();// "a5";
                                size = MeasureString(list_tb[ind].Text, list_tb[ind]);
                                list_tb[ind].Width = size.Width + 2;
                                list_tb[ind].Height = size.Height + 2;

                                if (i % 2 == 0)
                                {
                                    Canvas.SetLeft(list_tb[ind], (j + 0.5) * figure_width
                                        - list_tb[ind].Width / 2);
                                }
                                else
                                {
                                    Canvas.SetLeft(list_tb[ind], even_row_margin_left + (j + 0.5) * figure_width
                                    - list_tb[ind].Width / 2);
                                }
                            }
                            
                            ind++;
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < rows_nr; i++)
                {
                    for (int j = 0; j < cols_nr; j++)
                    {
                        if (list_tb[ind].Text != elements[ind].symbol)
                        {
                            list_tb[ind].Text = elements[ind].symbol;
                            //list_tb[ind].Text = random.Next(0, 100).ToString();// font_size.ToString();// "a5";
                            size = MeasureString(list_tb[ind].Text, list_tb[ind]);
                            list_tb[ind].Width = size.Width + 2;
                            list_tb[ind].Height = size.Height + 2;
                        }

                        ind++;
                    }
                }
            }   
        }

        Size MeasureString(string candidate, TextBlock textBlock)
        {
            var formattedText = new FormattedText(
                candidate,
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface(textBlock.FontFamily, textBlock.FontStyle, textBlock.FontWeight, textBlock.FontStretch),
                textBlock.FontSize,
                Brushes.Black,
                new NumberSubstitution(),
                1);

            return new Size(formattedText.Width, formattedText.Height);
        }

        private void CenterWindowOnScreen()
        {
            this.Left = this.Top = 0;
        }

        public int GetWindowsScaling()
        {
            return (int)(100 * System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width 
                / SystemParameters.PrimaryScreenWidth);
        }
    }

    public static class ExtensionMethods
    {
        private static readonly Action EmptyDelegate = delegate { };
        public static void Refresh(this UIElement uiElement)
        {
            uiElement.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate);
        }
    }
}