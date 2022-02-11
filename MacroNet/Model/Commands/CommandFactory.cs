using System.Windows.Forms;

namespace MacroNet.Model.Commands
{
    public static class CommandFactory
    {
        public static ICommand GetMouseCommand(MouseEventFlags mouseEventFlags, int x, int y)
        {
            return new MouseCommand(mouseEventFlags, x, y);
        }

        public static ICommand GetKeyboardCommand(KeyBoardEvent keyBoardEvent, Keys keys)
        {
            return new KeyBoardCommand(keyBoardEvent, keys);
        }

        public static ICommand GetDelayCommand(int delayMs)
        {
            return new DelayCommand(delayMs);
        }
    }
}
