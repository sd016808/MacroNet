using MaterialSkin.Controls;
using System;
using System.Windows.Forms;

namespace MacroNet
{
    public partial class InputKeyDialog : MaterialForm
    {
        public Keys Keys => (Keys)this.comboBox1.SelectedItem;
        public InputKeyDialog(Keys defaultKeys = Keys.F1)
        {
            InitializeComponent();
            this.comboBox1.DataSource = Enum.GetValues(typeof(Keys));
            this.comboBox1.SelectedItem = defaultKeys;
            this.StartPosition = FormStartPosition.CenterParent;
        }
    }
}
