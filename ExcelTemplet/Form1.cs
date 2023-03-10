using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelTemplet.BaseMethod;

namespace ExcelTemplet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataRows DS = new DataRows();
        List<object[]> OBJ = new List<object[]>();
        private void Form1_Load(object sender, EventArgs e)
        {
            DS.Columns = new string[] { "序号", "姓名", "年龄", "得分", "等级" };
            DS.type = new object[] { typeof(int), typeof(string), typeof(int), typeof(double), typeof(int) };
            OBJ.Add(new object[] { 1, "张格拉", 23, 65.34, 5 });
            OBJ.Add(new object[] { 2, "李冰", 23, 98.6, 1 });
            OBJ.Add(new object[] { 3, "张梅", 22, 89.2, 2 });
            OBJ.Add(new object[] { 4, "周遥", 23, 87.4, 2 });
            OBJ.Add(new object[] { 5, "古思优", 24, 99.8, 1 });
            dataGridView1.DataSource = BuidData(DS, OBJ);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog OP = new OpenFileDialog();
            OP.Filter = "jpg|*.jpg";
            if (OP.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.Image = Image.FromFile(OP.FileName);
            }
        }
        string File = string.Empty;

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog OP = new OpenFileDialog();
            OP.Filter = "xlsx|*.xlsx";
            if (OP.ShowDialog() == DialogResult.OK)
            {
                File = OP.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(File))
            {
                MessageBox.Show("请先选择一个模版");
                return;
            }
              if (pictureBox1.Image==null)
            {
                MessageBox.Show("请先选择一个图片");
                return;
            }
            SaveFileDialog SP = new SaveFileDialog();
            SP.Filter = "xlsx|*.xlsx";
            if (SP.ShowDialog() == DialogResult.OK)
            {
                Avant_ExcelTempletHelper.ExcelTemplet excel = new Avant_ExcelTempletHelper.ExcelTemplet(SP.FileName, File);
                DataTable dt = BuidData(DS, OBJ);
               excel.SetDataTableToExcelTemplet(dt, "成绩表");
                excel.SetStringToExcelTemplet("变量1", textBox1.Text);
                excel.SetStringToExcelTemplet("变量2", textBox2.Text);
                excel.SetStringToExcelTemplet("变量3", textBox3.Text);
                excel.SetStringToExcelTemplet("变量4", textBox4.Text);
                excel.SetStringToExcelTemplet("变量5", textBox5.Text);
                Image image = pictureBox1.Image;
                MemoryStream ms = new MemoryStream();
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                Stream stream = ms;
                //excel.PutImage(stream, 4, 0, 30, 10);
                 //excel.PutImage(stream, 1, 8, 3, 9);  //行参数间隔需要超过2行，
                excel.PutImage(stream,Convert.ToInt32(numericUpDown1.Value) , Convert.ToInt32(numericUpDown2.Value), Convert.ToInt32(numericUpDown3.Value), Convert.ToInt32(numericUpDown4.Value));

                excel.SavePath = SP.FileName;

                var tempback = excel.SafeSave();
                if (tempback == "OK")
                {
                    MessageBox.Show("保存完成");
                }
                else MessageBox.Show("出现错误" + tempback);

                //try
                //{
                //    excel.Save();
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("错误:" + ex.Message.ToString());
                //    // throw;
                //}
            }        
        }

        private void button4_Click(object sender, EventArgs e)
        {
            numericUpDown1.Value = 4;
            numericUpDown2.Value = 0;
            numericUpDown3.Value = 31;
            numericUpDown4.Value = 10;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            numericUpDown1.Value = 1;
            numericUpDown2.Value = 8;
            numericUpDown3.Value = 3;
            numericUpDown4.Value = 9;
        }
    }
}
