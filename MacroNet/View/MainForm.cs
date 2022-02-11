using Gma.System.MouseKeyHook;
using MacroNet.Model;
using MacroNet.Model.Commands;
using MaterialSkin;
using MaterialSkin.Controls;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MacroNet
{
    public partial class MainForm : MaterialForm
    {
        #region Field
        /// <summary>
        /// 使用者設定參數
        /// </summary>
        private Profile _userSetting = new Profile();

        /// <summary>
        /// 正在編輯的使用者設定檔
        /// </summary>
        private UC_config _current_Edit_Config = null;

        /// <summary>
        /// 是否正在顯示左側油圖 (為了做Sliding的特效才使用)
        /// </summary>
        private bool _isShow = true;

        /// <summary>
        /// 使用者設定參數存放路徑
        /// </summary>
        private string _setting_path = "setting.json";

        /// <summary>
        /// Global Hook
        /// </summary>
        private IKeyboardMouseEvents _globalHook;
        #endregion

        #region Constructor
        /// <summary>
        /// Constructor
        /// </summary>
        public MainForm()
        {
            InitializeComponent();

            // Skin
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);

            // Load Setting
            LoadUserSetting();

            // Initialize
            InitForm();

            // Hook
            _globalHook = Hook.GlobalEvents();
            _globalHook.KeyDown += OnKeyDown;
        }

        /// <summary>
        /// 初始化 UI 設定
        /// </summary>
        private void InitForm()
        {
            listView_macros.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            listView_macros.View = View.Details;
            listview_commands.View = View.Details;

            listView_macros.SizeChanged += ListView_SizeChanged;
            listview_commands.SizeChanged += ListView_SizeChanged;

            List<UC_config> ucConfigs = new List<UC_config>();
            for (int i = 0; i < _userSetting.Configs.Count; i++)
            {
                ucConfigs.Add(new UC_config(_userSetting.Configs[i]) { Dock = DockStyle.Fill });
                tableLayoutPanel5.Controls.Add(ucConfigs[i], i + 1, 0);
                ucConfigs[i].ConfigName = $"預設 {i + 1}";
                ucConfigs[i].On_Clicked += On_Config_Changed;

            }

            ucConfigs[0].Active();
            _current_Edit_Config = ucConfigs[0];

            tableLayoutPanel6.Visible = false;
            listview_commands.Visible = false;
            comboBox1.DataSource = Enum.GetValues(typeof(Keys));
            On_Config_Changed(ucConfigs[0]);

            /// <summary>
            /// 設定檔切換
            /// </summary>
            /// <param name="config"></param>
            void On_Config_Changed(UC_config config)
            {
                _current_Edit_Config = config;

                // 清除光標
                ucConfigs.Except(new[] { config }).ToList().ForEach(x => x.InActive());

                // 重新載入Macro清單
                listView_macros.Items.Clear();
                foreach (var m in _current_Edit_Config.Config.Macros)
                {
                    ListViewItem item = new ListViewItem();
                    item.Text = m.Name;
                    item.Tag = m;
                    item.Checked = m.Enabled;

                    listView_macros.Items.Add(item);
                }

                listView_macros.View = View.Details;
                listView_macros_SelectedIndexChanged(listView_macros, null);
            }
        }
        #endregion

        #region Save/Load
        /// <summary>
        /// 載入使用者參數設定
        /// </summary>
        private void LoadUserSetting()
        {
            try
            {
                if (File.Exists(_setting_path))
                {
                    var settings = new JsonSerializerSettings
                    {
                        PreserveReferencesHandling = PreserveReferencesHandling.Objects,
                        ReferenceLoopHandling = ReferenceLoopHandling.Serialize,
                        NullValueHandling = NullValueHandling.Ignore,
                        TypeNameHandling = TypeNameHandling.Auto
                    };


                    var jsonText = File.ReadAllText(_setting_path);
                    _userSetting = JsonConvert.DeserializeObject<Profile>(jsonText, settings);
                }
                else
                {
                    _userSetting = new Profile();
                    _userSetting.Init();
                }
            }
            catch
            {
                _userSetting = new Profile();
                _userSetting.Init();
            }
        }

        /// <summary>
        /// 儲存使用者參數設定
        /// </summary>
        private void SaveUserSetting()
        {
            var settings = new JsonSerializerSettings
            {
                PreserveReferencesHandling = PreserveReferencesHandling.Objects,
                ReferenceLoopHandling = ReferenceLoopHandling.Serialize,
                NullValueHandling = NullValueHandling.Ignore,
                TypeNameHandling = TypeNameHandling.Auto
            };

            string output = JsonConvert.SerializeObject(_userSetting, settings);
            File.WriteAllText(_setting_path, output);
        }
        #endregion

        #region Timer
        /// <summary>
        /// Timer for 紀錄滑鼠座標
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, System.EventArgs e)
        {
            toolStripStatusLabel1.Text = $"滑鼠座標: {Cursor.Position.X},{Cursor.Position.Y}";
        }

        /// <summary>
        /// 展開/收集左側的油圖的 Timer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer2_Tick(object sender, System.EventArgs e)
        {
            if (_isShow)
            {
                if (flowLayoutPanel1.Width >= 240)
                {
                    timer2.Stop();
                    return;
                }

                flowLayoutPanel1.Width += 120;
            }
            else
            {
                if (flowLayoutPanel1.Width <= 0)
                {
                    flowLayoutPanel1.Hide();
                    timer2.Stop();
                    return;
                }

                flowLayoutPanel1.Width -= 120;
            }
        }
        #endregion

        #region Global Hook
        /// <summary>
        /// Global Hook for OnKeyDown
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">KeyEventArgs</param>
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (_current_Edit_Config == null)
                return;

            foreach (var macro in _current_Edit_Config.Config.Macros)
            {
                if (macro.Enabled && macro.Trigger == e.KeyCode)
                {
                    foreach (var command in macro.Commands)
                    {
                        command.Do();
                    }
                }
            }
        }
        #endregion

        #region Appearance
        /// <summary>
        /// 根據 ListView SizeChanged 同步調整 Column Width
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListView_SizeChanged(object sender, EventArgs e)
        {
            ListView listView = sender as ListView;
            listView.Columns[0].Width = listView.Width;
        }

        /// <summary>
        /// Draw Boarder for Config Setting Area
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tableLayoutPanel5_CellPaint(object sender, System.Windows.Forms.TableLayoutCellPaintEventArgs e)
        {
            var panel = sender as TableLayoutPanel;
            e.Graphics.SmoothingMode = SmoothingMode.HighQuality;
            var rectangle = e.CellBounds;
            using (var pen = new Pen(Color.Black, 1))
            {
                pen.Alignment = System.Drawing.Drawing2D.PenAlignment.Center;
                pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;

                if (e.Row == (panel.RowCount - 1))
                {
                    rectangle.Height -= 1;
                }

                if (e.Column == (panel.ColumnCount - 1))
                {
                    rectangle.Width -= 1;
                }

                e.Graphics.DrawRectangle(pen, rectangle);
            }
        }
        #endregion

        /// <summary>
        /// 新增巨集
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void materialButton4_Click(object sender, System.EventArgs e)
        {
            string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入巨集名稱", DefaultResponse: "巨集1");
            if (string.IsNullOrWhiteSpace(input))
                return;

            listView_macros.Items.Clear();
            var macro = new Macro(input);
            _current_Edit_Config.Config.Macros.Add(macro);
            foreach (var m in _current_Edit_Config.Config.Macros)
            {
                ListViewItem item = new ListViewItem();
                item.Text = m.Name;
                item.Tag = m;
                item.Checked = m.Enabled;

                listView_macros.Items.Add(item);
            }
        }

        /// <summary>
        /// 刪除全部巨集
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void materialButton1_Click(object sender, System.EventArgs e)
        {
            if (MessageBox.Show("確定要刪除全部的Macro嗎?", "確認", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            _current_Edit_Config.Config.Macros.Clear();
            listView_macros.Items.Clear();
            listView_macros_SelectedIndexChanged(listView_macros, null);
        }

        /// <summary>
        /// 啟用巨集或關閉巨集
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView_macros_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            (e.Item.Tag as Macro).Enabled = e.Item.Checked;
        }

        /// <summary>
        /// 啟動/關閉 紀錄滑鼠座標的 Timer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void materialCheckbox1_CheckedChanged(object sender, System.EventArgs e)
        {
            if (materialCheckbox1.Checked)
                timer1.Start();
            else
            {
                timer1.Stop();
                toolStripStatusLabel1.Text = "滑鼠座標: x, y";
            }
        }

        /// <summary>
        /// 展開/收集左側的油圖
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pictureBox3_Click(object sender, System.EventArgs e)
        {
            if (flowLayoutPanel1.Visible)
            {
                _isShow = false;
                timer2.Start();
            }
            else
            {
                _isShow = true;
                flowLayoutPanel1.Show();
                timer2.Start();
            }
        }

        /// <summary>
        /// 點選不同的 Macro
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView_macros_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (listView_macros.SelectedItems.Count == 0)
            {
                tableLayoutPanel6.Visible = false;
                listview_commands.Visible = false;
            }
            else
            {
                var macro = listView_macros.SelectedItems[0].Tag as Macro;
                tableLayoutPanel6.Visible = true;
                listview_commands.Visible = true;


                textBox1.Text = macro.Name;
                comboBox1.SelectedItem = macro.Trigger;
                listview_commands.Items.Clear();
                foreach (var command in macro.Commands)
                {
                    ListViewItem item = new ListViewItem();
                    item.Text = command.Text;
                    item.Tag = command;
                    listview_commands.Items.Add(item);
                }
            }
        }

        /// <summary>
        /// 修改 Macro 的名稱
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var macro = listView_macros.SelectedItems[0].Tag as Macro;
            macro.Name = textBox1.Text;
            listView_macros.SelectedItems[0].Text = macro.Name;

        }

        /// <summary>
        /// 將選擇的 Command 向上移動
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void materialButton2_Click(object sender, EventArgs e)
        {
            try
            {
                if (listview_commands.SelectedItems.Count > 0)
                {
                    ListViewItem selected = listview_commands.SelectedItems[0];
                    int indx = selected.Index;
                    int totl = listview_commands.Items.Count;

                    if (indx == 0)
                    {
                        listview_commands.Items.Remove(selected);
                        listview_commands.Items.Insert(totl - 1, selected);
                    }
                    else
                    {
                        listview_commands.Items.Remove(selected);
                        listview_commands.Items.Insert(indx - 1, selected);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// 將選擇的 Command 向下移動
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void materialButton3_Click(object sender, EventArgs e)
        {
            try
            {
                if (listview_commands.SelectedItems.Count > 0)
                {
                    ListViewItem selected = listview_commands.SelectedItems[0];
                    int indx = selected.Index;
                    int totl = listview_commands.Items.Count;

                    if (indx == totl - 1)
                    {
                        listview_commands.Items.Remove(selected);
                        listview_commands.Items.Insert(0, selected);
                    }
                    else
                    {
                        listview_commands.Items.Remove(selected);
                        listview_commands.Items.Insert(indx + 1, selected);
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        /// <summary>
        /// 新增滑鼠相關的命令
        /// </summary>
        /// <param name="mouseEventFlags">滑鼠的命令類型</param>
        /// <param name="x">滑鼠座標 X</param>
        /// <param name="y">滑鼠座標 Y</param>
        private void AddMouseCommand(MouseEventFlags mouseEventFlags, int x, int y)
        {
            var macro = GetCurrentMacro();
            var command = CommandFactory.GetMouseCommand(mouseEventFlags, x, y);
            macro.Commands.Add(command);
            ListViewItem item = new ListViewItem();
            item.Text = command.Text;
            item.Tag = command;
            listview_commands.Items.Add(item);
            listview_commands.Columns[0].Width = listview_commands.Width;
        }

        /// <summary>
        /// 新增鍵盤相關的命令
        /// </summary>
        /// <param name="keyBoardEvent">鍵盤的命令類型</param>
        /// <param name="key">指定按鍵碼 (Key Code)</param>
        private void AddKeyCommand(KeyBoardEvent keyBoardEvent, Keys key)
        {
            var macro = GetCurrentMacro();
            var command = CommandFactory.GetKeyboardCommand(keyBoardEvent, key);
            macro.Commands.Add(command);
            ListViewItem item = new ListViewItem();
            item.Text = command.Text;
            item.Tag = command;
            listview_commands.Items.Add(item);
            listview_commands.Columns[0].Width = listview_commands.Width;
        }

        /// <summary>
        /// 新增滑鼠左鍵按下的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void 左鍵按下ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入滑鼠座標 x,y", DefaultResponse: "0,0");
                int x = int.Parse(input.Split(',')[0]);
                int y = int.Parse(input.Split(',')[1]);

                AddMouseCommand(MouseEventFlags.LeftDown, x, y);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 新增滑鼠左鍵放開的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void 左鍵放開ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入滑鼠座標 x,y", DefaultResponse: "0,0");
                int x = int.Parse(input.Split(',')[0]);
                int y = int.Parse(input.Split(',')[1]);

                AddMouseCommand(MouseEventFlags.LeftUp, x, y);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 新增滑鼠右鍵按下的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void 右鍵按下ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入滑鼠座標 x,y", DefaultResponse: "0,0");
                int x = int.Parse(input.Split(',')[0]);
                int y = int.Parse(input.Split(',')[1]);

                AddMouseCommand(MouseEventFlags.RightDown, x, y);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 新增滑鼠右鍵放開的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void 右鍵放開ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入滑鼠座標 x,y", DefaultResponse: "0,0");
                int x = int.Parse(input.Split(',')[0]);
                int y = int.Parse(input.Split(',')[1]);

                AddMouseCommand(MouseEventFlags.RightUp, x, y);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 新增鍵盤按下的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void 鍵盤按下ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                InputKeyDialog inputKeyDialog = new InputKeyDialog();
                if (inputKeyDialog.ShowDialog() != DialogResult.OK)
                    return;

                AddKeyCommand(KeyBoardEvent.Key_Down, inputKeyDialog.Keys);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 新增鍵盤放開的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void 鍵盤放開ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                InputKeyDialog inputKeyDialog = new InputKeyDialog();
                if (inputKeyDialog.ShowDialog() != DialogResult.OK)
                    return;

                AddKeyCommand(KeyBoardEvent.Key_Up, inputKeyDialog.Keys);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 新增 Delay 的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void delayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入 delay 時間(ms)", DefaultResponse: "50");
                int delayMs = int.Parse(input.Split(',')[0]);

                var macro = GetCurrentMacro();
                var command = CommandFactory.GetDelayCommand(delayMs);
                macro.Commands.Add(command);
                ListViewItem item = new ListViewItem();
                item.Text = command.Text;
                item.Tag = command;
                listview_commands.Items.Add(item);
                listview_commands.Columns[0].Width = listview_commands.Width;
            }
            catch
            {
            }
        }

        /// <summary>
        /// Form Closing Event
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">FormClosingEventArgs</param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveUserSetting();
        }

        /// <summary>
        /// 刪除目前Focus的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listview_commands.SelectedItems.Count <= 0)
                return;

            var command = listview_commands.SelectedItems[0].Tag as ICommand;
            var macro = GetCurrentMacro();
            if (macro == null)
                return;

            macro.Commands.Remove(command);
            listview_commands.Items.Remove(listview_commands.SelectedItems[0]);
        }

        /// <summary>
        /// 取得正在編輯的 Macro
        /// </summary>
        /// <returns></returns>
        private Macro GetCurrentMacro()
        {
            if (_current_Edit_Config == null)
                return null;

            if (listView_macros.SelectedItems.Count <= 0)
                return null;

            return listView_macros.SelectedItems[0].Tag as Macro;
        }

        /// <summary>
        /// 新增滑鼠左鍵Click一次的命令
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 左鍵點擊ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入滑鼠座標 x,y", DefaultResponse: "0,0");
                int x = int.Parse(input.Split(',')[0]);
                int y = int.Parse(input.Split(',')[1]);

                AddMouseCommand(MouseEventFlags.LeftClick, x, y);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 新增滑鼠右鍵Click一次的命令
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void 右鍵點擊ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("請輸入滑鼠座標 x,y", DefaultResponse: "0,0");
                int x = int.Parse(input.Split(',')[0]);
                int y = int.Parse(input.Split(',')[1]);

                AddMouseCommand(MouseEventFlags.RightClick, x, y);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 設定正在編輯的Macro的觸發鍵
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">EventArgs</param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var macro = GetCurrentMacro();
            if (macro == null)
                return;

            macro.Trigger = (Keys)comboBox1.SelectedItem;
        }

        /// <summary>
        /// 修改目前的指令參數
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">MouseEventArgs</param>
        private void listView_commands_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listview_commands.SelectedItems.Count <= 0)
                return;

            var command = listview_commands.SelectedItems[0].Tag as ICommand;
            command.Edit();
            listview_commands.SelectedItems[0].Text = command.Text;
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var macro = GetCurrentMacro();
            if (macro == null)
                return;

            _current_Edit_Config.Config.Macros.Remove(macro);
            listView_macros.Items.Remove(listView_macros.SelectedItems[0]);
        }

        private void contextMenuStrip2_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var macro = GetCurrentMacro();
            if (macro == null)
                e.Cancel = true;
        }
    }
}
