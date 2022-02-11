using System.Collections.Generic;

namespace MacroNet.Model
{
    public class Config
    {
        /// <summary>
        /// Config Name
        /// </summary>
        public string Name { get; set; }
        public List<Macro> Macros { get; set; } = new List<Macro>();
        public Config()
        {
        }
    }
}
