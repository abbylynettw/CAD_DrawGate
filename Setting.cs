using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MultiExcelMultiDoor
{
    public partial class SettingForm : Form
    {
        public SettingForm()
        {
            InitializeComponent();
            SetValue();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var file = IOHelper.GetFile();
            if (!string.IsNullOrEmpty(file))
            {
                INIHelper.IniWriteValue("FOlDERPATH", "DWGPath", file);
                SetValue();
                MessageBox.Show("修改成功！");
            }
        }
        
        private void SetValue()
        {
            this.textBox1.Text = INIHelper.IniReadValue("FOlDERPATH", "DWGPath", "无");
            this.textBox2.Text = INIHelper.IniReadValue("FOlDERPATH", "DXFPath", "无");
            this.textBox3.Text = INIHelper.IniReadValue("FOlDERPATH", "EXCELCornerPath", "无");
            this.textBox4.Text = INIHelper.IniReadValue("FOlDERPATH", "EXCELPolyPath", "无");
            this.textBox5.Text = INIHelper.IniReadValue("FILEPATH", "DWGCornerPath", "无");
            this.textBox6.Text = INIHelper.IniReadValue("FILEPATH", "DWGBlOCKPATH", "无");
            this.textBox7.Text = INIHelper.IniReadValue("LAYERNAME", "layer", "无");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var file = IOHelper.GetFile();
            if (!string.IsNullOrEmpty(file))
            {
                INIHelper.IniWriteValue("FOlDERPATH", "DXFPath", file);
                SetValue();
                MessageBox.Show("修改成功！");
            }

         
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var file = IOHelper.GetFile();
            if (!string.IsNullOrEmpty(file))
            {
                INIHelper.IniWriteValue("FOlDERPATH", "EXCELCornerPath", file);
                SetValue();
                MessageBox.Show("修改成功！");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var file = IOHelper.GetFile();
            if (!string.IsNullOrEmpty(file))
            {
                INIHelper.IniWriteValue("FOlDERPATH", "EXCELPolyPath", file);
                SetValue();
                MessageBox.Show("修改成功！");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var file = IOHelper.SelectPath();
            if (!string.IsNullOrEmpty(file))
            {
                INIHelper.IniWriteValue("FILEPATH", "DWGCornerPath", file);
                SetValue();
                MessageBox.Show("修改成功！");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var file = IOHelper.SelectPath();
            if (!string.IsNullOrEmpty(file))
            {
                INIHelper.IniWriteValue("FILEPATH", "DWGBlOCKPATH", file);
                SetValue();
                MessageBox.Show("修改成功！");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var text = this.textBox7.Text;
            if (!string.IsNullOrEmpty(text))
            {
                INIHelper.IniWriteValue("LAYERNAME", "layer", text);
                SetValue();
                MessageBox.Show("修改成功！");
            }
        }

    }
}
