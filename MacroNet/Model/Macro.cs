using MacroNet.Model.Commands;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MacroNet.Model
{
    public class Macro
    {
        /// <summary>
        /// Macro Name
        /// </summary>
        public string Name { get; set; }
        public List<ICommand> Commands { get; set; } = new List<ICommand>();
        public bool Enabled { get; set; }
        public Keys Trigger { get; set; }
        public Macro(string name, bool enabled = true, Keys keys = Keys.F1)
        {
            Name = name;
            Enabled = enabled;
            Trigger = keys;
        }
    }
}
