using System;
using System.Runtime.InteropServices;

namespace MacroNet.Model.Commands
{
    public class MouseCommand : ICommand
    {
        public CommandType CommandType => CommandType.Mouse;
        public MouseEventFlags Event { get; set; }
        public MousePoint Position { get; set; }
        public string Text
        {
            get
            {
                string text = string.Empty;
                switch (Event)
                {
                    case MouseEventFlags.LeftDown:
                        text += "左鍵按下 ";
                        break;
                    case MouseEventFlags.LeftUp:
                        text += "左鍵放開 ";
                        break;
                    case MouseEventFlags.LeftClick:
                        text += "左鍵點擊 ";
                        break;
                    case MouseEventFlags.RightDown:
                        text += "右鍵按下 ";
                        break;
                    case MouseEventFlags.RightUp:
                        text += "右鍵放開 ";
                        break;
                    case MouseEventFlags.RightClick:
                        text += "右鍵點擊 ";
                        break;
                    default:
                        throw new NotImplementedException();
                }

                return text + $" {Position.X},{Position.Y}";
            }
        }

        public MouseCommand(MouseEventFlags mouseEventFlags, int x, int y)
        {
            Event = mouseEventFlags;
            Position = new MousePoint(x, y);
        }

        public void Do()
        {
            SetCursorPosition(Position);
            MouseEvent(Event);
        }

        public void Edit()
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入滑鼠座標 x,y", DefaultResponse: $"{Position.X},{Position.Y}");
                int x = int.Parse(input.Split(',')[0]);
                int y = int.Parse(input.Split(',')[1]);

                Position = new MousePoint(x, y);
            }
            catch
            {
            }
        }

        private void SetCursorPosition(MousePoint point)
        {
            SetCursorPos(point.X, point.Y);
        }

        private void MouseEvent(MouseEventFlags value)
        {
            MousePoint position = GetCursorPosition();

            if (Event == MouseEventFlags.LeftClick)
            {
                mouse_event
                ((int)MouseEventFlags.LeftDown,
                    position.X,
                    position.Y,
                    0,
                    0);
                mouse_event
                ((int)MouseEventFlags.LeftUp,
                    position.X,
                    position.Y,
                    0,
                    0);
            }
            else if (Event == MouseEventFlags.RightClick)
            {
                mouse_event
                ((int)MouseEventFlags.RightDown,
                    position.X,
                    position.Y,
                    0,
                    0);
                mouse_event
                ((int)MouseEventFlags.RightUp,
                    position.X,
                    position.Y,
                    0,
                    0);
            }
            else
            {
                mouse_event
                ((int)(value),
                 position.X,
                 position.Y,
                 0,
                 0);
            }
        }

        private MousePoint GetCursorPosition()
        {
            MousePoint currentMousePoint;
            var gotPoint = GetCursorPos(out currentMousePoint);
            if (!gotPoint) { currentMousePoint = new MousePoint(0, 0); }
            return currentMousePoint;
        }

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out MousePoint lpMousePoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [StructLayout(LayoutKind.Sequential)]
        public struct MousePoint
        {
            public int X;
            public int Y;

            public MousePoint(int x, int y)
            {
                X = x;
                Y = y;
            }
        }
    }
}
