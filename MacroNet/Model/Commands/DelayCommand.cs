using System.Threading;

namespace MacroNet.Model.Commands
{
    public class DelayCommand : ICommand
    {
        public CommandType CommandType => CommandType.Delay;
        public int DelayMS { get; set; }
        public string Text => $"Delay {DelayMS}ms";

        public DelayCommand(int delayMS)
        {
            DelayMS = delayMS;
        }

        public void Do()
        {
            Thread.Sleep(DelayMS);
        }

        public void Edit()
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入 delay 時間(ms)", DefaultResponse: DelayMS.ToString());
                DelayMS = int.Parse(input.Split(',')[0]);
            }
            catch
            {
            }
        }
    }
}
