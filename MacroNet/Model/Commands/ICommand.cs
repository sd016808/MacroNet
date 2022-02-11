namespace MacroNet.Model.Commands
{
    public interface ICommand
    {
        CommandType CommandType { get; }
        string Text { get; }
        void Do();
        void Edit();
    }
}
