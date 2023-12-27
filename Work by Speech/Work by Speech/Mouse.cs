using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using WindowsInput.Native;

namespace Speech
{
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetCursorPos(out Point lpPoint);

        void left_down()
        {
            sim.Mouse.LeftButtonDown();
        }

        void left_up()
        {
            sim.Mouse.LeftButtonUp();
        }

        void right_down()
        {
            sim.Mouse.RightButtonDown();
        }

        void right_up()
        {
            sim.Mouse.RightButtonUp();
        }
        
        public void move_mouse(int x, int y)
        {
            int scaling = GetWindowsScaling();
            x = (int)(x * scaling / 100);
            y = (int)(y * scaling / 100);

            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(x, y);
            //SetCursorPos(x, y);
        }

        void move_mouse_by(int x, int y)
        {
            x += System.Windows.Forms.Cursor.Position.X;
            y += System.Windows.Forms.Cursor.Position.Y;

            if (x < 0) x = 0;
            else if (x > screen_width - 1) x = screen_width - 1;

            if (y < 0) y = 0;
            else if (y > screen_height - 1) y = screen_height - 1;

            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(x, y);
            //SetCursorPos(x, y);
        }

        void real_mouse_move(int x, int y, bool max_speed = false)
        {
            int scaling = GetWindowsScaling();
            x = (int)(x * scaling / 100);
            y = (int)(y * scaling / 100);

            int new_x = x;
            int new_y = y;
            int current_x = System.Windows.Forms.Cursor.Position.X;
            int current_y = System.Windows.Forms.Cursor.Position.Y;

            int movement_speed = 100; //it's much faster when run from visual studio

            if (screen_height >= 2160)
                movement_speed = 200;
            else if (screen_height >= 4320)
                movement_speed = 400;
            else if (screen_height >= 8640)
                movement_speed = 800;

            int i = 1;

            while (current_x != new_x || current_y != new_y)
            {
                if (current_x < new_x) current_x++;
                else if (current_x > new_x) current_x--;

                if (current_y < new_y) current_y++;
                else if (current_y > new_y) current_y--;

                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(current_x, current_y);

                if (max_speed == false && i % movement_speed == 0)
                {
                    real_sleep(1);
                }
                i++;
            }
        }

        void real_move_mouse_by(int x, int y, bool max_speed = false)
        {
            int current_x = System.Windows.Forms.Cursor.Position.X;
            int current_y = System.Windows.Forms.Cursor.Position.Y;

            int new_x = x + current_x;
            int new_y = y + current_y;

            if (new_x < 0) new_x = 0;
            else if (new_x > screen_width - 1) new_x = screen_width - 1;

            if (new_y < 0) new_y = 0;
            else if (new_y > screen_height - 1) new_y = screen_height - 1;

            int movement_speed = 100; //it's much faster when run from visual studio

            if (screen_height >= 2160)
                movement_speed = 200;
            else if (screen_height >= 4320)
                movement_speed = 400;
            else if (screen_height >= 8640)
                movement_speed = 800;

            int i = 1;

            while (current_x != new_x || current_y != new_y)
            {
                if (current_x < new_x) current_x++;
                else if (current_x > new_x) current_x--;

                if (current_y < new_y) current_y++;
                else if (current_y > new_y) current_y--;

                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(current_x, current_y);

                if (max_speed == false && i % movement_speed == 0)
                {
                    real_sleep(1);
                }
                i++;
            }
        }

        void LMBClick(int x = -1, int y = -1, int time = 75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            //user may forget that right button is pressed or press it by mistake without noticing
            //(holding RMB prevents LMB clicking)
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON))
            {
                right_up();
            }
            freeze_mouse(x, y, 10);
            left_down();
            freeze_mouse(x, y, time);
            left_up();
            freeze_mouse(x, y, 10);
        }

        void RMBClick(int x = -1, int y = -1, int time = 75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            freeze_mouse(x, y, 10);
            right_down();
            freeze_mouse(x, y, time);
            right_up();
            freeze_mouse(x, y, 10);
        }

        void DLMBClick(int x = -1, int y = -1, int time = 75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            //user may forget that right button is pressed or press it by mistake without noticing
            //(holding RMB prevents LMB clicking)
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON))
            {
                right_up();
            }
            LMBClick(x, y, time);
            LMBClick(x, y, time);
        }

        void TLMBClick(int x = -1, int y = -1, int time = 75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            //user may forget that right button is pressed or press it by mistake without noticing
            //(holding RMB prevents LMB clicking)
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON))
            {
                right_up();
            }
            LMBClick(x, y, time);
            LMBClick(x, y, time);
            LMBClick(x, y, time);
        }

        void LMBHold_toggle(int x = -1, int y = -1, int time = 75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            freeze_mouse(x, y, 10);
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.LBUTTON)==false)
            {
                left_down();
            }
            else
            {
                left_up();
            }
            freeze_mouse(x, y, time);
        }

        void RMBHold_toggle(int x = -1, int y = -1, int time = 75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            freeze_mouse(x, y, 10);
            if (sim.InputDeviceState.IsKeyDown(VirtualKeyCode.RBUTTON)==false)
            {
                right_down();
            }
            else
            {
                right_up();
            }
            freeze_mouse(x, y, time);
        }

        void LMBHold(int x = -1, int y = -1, int time = 75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            freeze_mouse(x, y, 10);
            left_down();             
            freeze_mouse(x, y, time);
        }

        void RMBHold(int x = -1, int y = -1, int time = 75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            freeze_mouse(x, y, 10);
            right_down();
            freeze_mouse(x, y, time);
        }

        void real_sleep(int time)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            do
            {
                Thread.Sleep(1);
            }
            while (stopwatch.ElapsedMilliseconds < time);
            stopwatch.Stop();
        }

        void freeze_mouse(int x = -1, int y = -1, int time=75)
        {
            if (x == -1 || y == -1)
            {
                x = System.Windows.Forms.Cursor.Position.X;
                y = System.Windows.Forms.Cursor.Position.Y;
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            do
            {
                move_mouse(x, y);
                Thread.Sleep(1);
            }
            while (stopwatch.ElapsedMilliseconds < time);
        }
    }
}