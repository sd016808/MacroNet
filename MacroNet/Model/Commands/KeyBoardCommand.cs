using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MacroNet.Model.Commands
{
    public class KeyBoardCommand : ICommand
    {
        public CommandType CommandType => CommandType.Keyboard;
        public KeyBoardEvent Event { get; set; }
        public Keys Keys { get; set; }
        public string Text
        {
            get
            {
                switch (Event)
                {
                    case KeyBoardEvent.Key_Down:
                        return $"按鍵按下 {Keys}";
                    case KeyBoardEvent.Key_Up:
                        return $"按鍵放開 {Keys}";
                    default:
                        throw new NotImplementedException();
                }
            }
        }

        public KeyBoardCommand(KeyBoardEvent keyBoardEvent, Keys keys)
        {
            Event = keyBoardEvent;
            Keys = Keys;
        }

        public void Do()
        {
            if (Event == KeyBoardEvent.Key_Down)
            {
                keybd_event((byte)(Keys),
                    0,
                    0,
                    0);
            }
            else
            {
                // Simulate a key release
                keybd_event((byte)(Keys),
                                 0,
                                 0x0002,
                                 0);
            }
        }

        public void Edit()
        {
            try
            {
                InputKeyDialog inputKeyDialog = new InputKeyDialog(Keys);
                if (inputKeyDialog.ShowDialog() != DialogResult.OK)
                    return;

                Keys = inputKeyDialog.Keys;
            }
            catch
            {
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
    }
}
