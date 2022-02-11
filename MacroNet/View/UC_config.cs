using MacroNet.Model;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MacroNet
{
    public partial class UC_config : UserControl
    {
        public Action<UC_config> On_Clicked;
        public Config Config { get; }
        public UC_config(Config config)
        {
            InitializeComponent();
            Config = config;
            InActive();
        }

        public string ConfigName
        {
            get => Config.Name;
            set
            {
                Config.Name = value;
                label1.Text = Config.Name;
            }
        }

        public void Active()
        {
            panel1.BackColor = Color.Purple;
        }

        public void InActive()
        {
            panel1.BackColor = Color.Transparent;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            On_Clicked?.Invoke(this);
            Active();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            On_Clicked?.Invoke(this);
            Active();
        }

        private void uc_config_Click(object sender, EventArgs e)
        {
            On_Clicked?.Invoke(this);
            Active();
        }

        private void tableLayoutPanel1_Click(object sender, EventArgs e)
        {
            On_Clicked?.Invoke(this);
            Active();
        }
    }
}
