using System.Collections.Generic;

namespace MacroNet.Model
{
    public class Profile
    {
        public static readonly int CONFIG_MAX_LIMIT = 6;
        public List<Config> Configs { get; set; } = new List<Config>();
        public Profile()
        {
        }

        public void Init()
        {
            for (int i = 0; i < CONFIG_MAX_LIMIT; i++)
            {
                Configs.Add(new Config());
            }
        }
    }
}
