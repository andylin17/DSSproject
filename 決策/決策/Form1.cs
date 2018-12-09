using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.IO;

namespace 決策
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            DBclass(getclass);
            Gridview();
        }
        string getsubject = @"select 課名 
                              from course 
                              where 學期 = ";
        string getclass = @"select 教室 
                            from classroom";
        string gettime = @"select 限制 
                            from prefer
                            where 教授 = ";
        string getlevel = @"select 等級 
                            from course
                            where 課名  = ";
        static List<string> temp = new List<string>();
        static List<string> isinsert = new List<string>();//課程是否輸進課表
        static List<string> coursearray = new List<string>();
        static List<string> classarray = new List<string>();
        //資料庫連線
        //----------------------------------------------
        private void DBclass(string str)
        {
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=Database2.mdb");
            conn.Open();
            OleDbCommand cmdforclass = new OleDbCommand(str, conn);
            OleDbDataReader classrd = cmdforclass.ExecuteReader();
            
            while (classrd.Read())
            {
                comboBox2.Items.Add(classrd.GetString(0));
                classarray.Add(classrd.GetString(0));
            }
            classrd.Close();
            conn.Close();
        }
        private void DBsubject(string str)
        {
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=Database2.mdb");
            conn.Open();
            OleDbCommand cmdforsubject = new OleDbCommand(str, conn);
            OleDbDataReader subjectrd = cmdforsubject.ExecuteReader();
            while (subjectrd.Read())
            {
                comboBox1.Items.Add(subjectrd.GetString(0));
                coursearray.Add(subjectrd.GetString(0));
            }
            subjectrd.Close();
            conn.Close();
        }
        private string DBgrade(string str)
        {
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=Database2.mdb");
            conn.Open();
            OleDbCommand cmdforgrade = new OleDbCommand(str, conn);
            OleDbDataReader graderd = cmdforgrade.ExecuteReader();
            graderd.Read();
            string grade = graderd.GetString(0);
            graderd.Close();
            conn.Close();

            return grade;
        }
        private List<string> DBtime(string str)
        {
            List<string> teacherlist = new List<string>();
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=Database2.mdb");
            conn.Open();
            OleDbCommand cmdfortime = new OleDbCommand(str, conn);
            OleDbDataReader timerd = cmdfortime.ExecuteReader();
            while (timerd.Read())
            {
                teacherlist.Add(timerd.GetString(0));
            }
            timerd.Close();
            conn.Close();

            return teacherlist;
        }
        private int DBlevel(string str)
        {
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=Database2.mdb");
            conn.Open();
            OleDbCommand cmdforlevel = new OleDbCommand(str, conn);
            OleDbDataReader levelrd = cmdforlevel.ExecuteReader();
            levelrd.Read();
            int level = levelrd.GetInt32(0);
            levelrd.Close();
            conn.Close();
            return level;
        }
        //----------------------------------------------
        private void Gridview()
        {
            DataGridViewRowCollection rows1 = dataGridView1.Rows;
            for (int i = 0; i <= 8; i++)
            { rows1.Add(new Object[] { null, null, null, null, null }); }
            DataGridViewRowCollection rows2 = dataGridView2.Rows;
            for (int i = 0; i <= 8; i++)
            { rows2.Add(new Object[] { null, null, null, null, null }); }
            DataGridViewRowCollection rows3 = dataGridView3.Rows;
            for (int i = 0; i <= 8; i++)
            { rows3.Add(new Object[] { null, null, null, null, null }); }
            DataGridViewRowCollection rows4 = dataGridView4.Rows;
            for (int i = 0; i <= 8; i++)
            { rows4.Add(new Object[] { null, null, null, null, null }); }
            DataGridViewRowCollection rows5 = dataGridView5.Rows;
            for (int i = 0; i <= 8; i++)
            { rows5.Add(new Object[] { null, null, null, null, null }); }
            DataGridViewRowCollection rows6 = dataGridView6.Rows;
            for (int i = 0; i <= 8; i++)
            { rows6.Add(new Object[] { null, null, null, null, null }); }
        }
        //---------------------------------------------------------------------------
        //設置Gridview  Row  Index
        //---------------------------------------------------------------------------
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dataGridView1.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dataGridView1.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dataGridView1.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dataGridView2.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dataGridView2.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dataGridView2.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        private void dataGridView3_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dataGridView3.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dataGridView3.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dataGridView3.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        private void dataGridView4_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dataGridView4.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dataGridView4.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dataGridView4.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        private void dataGridView5_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dataGridView5.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dataGridView5.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dataGridView5.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        //---------------------------------------------------------------------------
        //---------------------------------------------------------------------------
        //重設時間選單
        private void Timeinit()
        {
            for (int i = 0; i < 50; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
            
        }
        //重設分析結果顏色
        private void Colorinit()
        {
            for (int column = 0; column < 5; column++)
            {
                for (int row = 0; row < 10; row++)
                {
                    dataGridView5[column, row].Style.BackColor = Color.White;
                }
            }

        }
        //初始化課表
        private void GridViewinit(DataGridView gridView)
        {
            for (int Column = 0; Column < 5; Column++)
            {
                for (int Row = 0; Row < 10; Row++)
                {
                    gridView[Column, Row].Value = null;
                }
            }
            isinsert.Clear();
        }
        private int Level_Compare(int a,int b)
        {
            if (a < b)
                return b;
            else if (a > b)
                return a;
            else
                return a;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
            getsubject += Global.semester;
            DBsubject(getsubject);
            if(Global.semester.Contains("上"))
            {
                this.Text = "上 學期";
            }
            else
                this.Text = "下 學期";
        }
        
        //顯示授課老師
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string getteacher = "select 教授 from course where 課名=\"";
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=Database2.mdb");
            conn.Open();
            getteacher += comboBox1.SelectedItem.ToString() + "\"";
            OleDbCommand cmdforteacher = new OleDbCommand(getteacher, conn);
            OleDbDataReader teacherrd = cmdforteacher.ExecuteReader();
            if(teacherrd.Read())
            {
                label11.Text = teacherrd.GetString(0);
                label11.Visible = true;
            }
            string grade = "select 年級 from course where 課名 =\"" + comboBox1.SelectedItem + "\"";
            label13.Text = DBgrade(grade);
            label13.Visible = true;
        }
        private bool Intogridview(DataGridView dataGridView,bool inserttrue)
        {
            string message = "";
            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                string roomstr = checkedListBox1.CheckedItems[i].ToString();
                string weekstr = roomstr.Substring(0, 1);
                string timestr = roomstr.Substring(1);
                int weeknum = int.Parse(weekstr) - 1;
                int timenum = int.Parse(timestr) - 1;
                bool classroom = false;
                if (isinsert.Count == 0)
                {
                    classroom = true;
                }
                else if (dataGridView6[weeknum, timenum].Value == null && isinsert.Count != 0)
                {
                    classroom = true;
                }
                else if (dataGridView6[weeknum, timenum].Value != null && isinsert.Count != 0)
                {
                    if (!dataGridView6[weeknum, timenum].Value.ToString().Contains(comboBox2.SelectedItem.ToString()))
                    {
                        classroom = true;
                    }
                }
                DataGridViewRowCollection rows = dataGridView.Rows;
                if (dataGridView[weeknum, timenum].Value == null)
                {
                    if (isinsert.Contains(comboBox1.SelectedItem.ToString()) == false && classroom == true)
                    {
                        dataGridView[weeknum, timenum].Value = comboBox1.SelectedItem.ToString() + " " + comboBox2.SelectedItem.ToString();
                        dataGridView6[weeknum, timenum].Value += comboBox2.SelectedItem.ToString() + ",";
                        inserttrue = true;
                    }
                    else if (!classroom && !inserttrue)
                    {
                        message += roomstr + " 時段 " + comboBox2.SelectedItem.ToString() + " 已使用\n";
                    }

                }
                else
                {
                    message += roomstr + "時段已有課程 \n";
                }
            }
            if (message != "")
            {
                MessageBox.Show(message);
            }
            return inserttrue;
        }
        //輸入至gridview
        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null && comboBox2.SelectedItem != null && checkedListBox1.CheckedItems.Count != 0)
            {
                string grade = "select 年級 from course where 課名 =\"" + comboBox1.SelectedItem + "\"";
                bool inserttrue = false;

                switch (DBgrade(grade))//判斷輸入至幾年級課表
                {
                    case "一":
                        inserttrue = Intogridview(dataGridView1, inserttrue);
                        break;
                    case "二":
                        inserttrue = Intogridview(dataGridView2, inserttrue);
                        break;
                    case "三":
                        inserttrue = Intogridview(dataGridView3, inserttrue);
                        break;
                    case "四":
                        inserttrue = Intogridview(dataGridView4, inserttrue);
                        break;
                }
                //取消時間勾選
                Timeinit();
                //
                if (isinsert.Contains(comboBox1.SelectedItem.ToString()) == false
                    && inserttrue == true)
                {
                    isinsert.Add(comboBox1.SelectedItem.ToString());
                }
                else if (isinsert.Contains(comboBox1.SelectedItem.ToString()) == true)
                {
                    MessageBox.Show("課表裡已有 " + comboBox1.SelectedItem.ToString());
                }

            }
            else if (comboBox1.SelectedItem == null && comboBox2.SelectedItem == null && checkedListBox1.CheckedItems.Count == 0)
                MessageBox.Show("請輸入 課程 與 教室 與 時間","",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            else if (comboBox1.SelectedItem != null && comboBox2.SelectedItem == null && checkedListBox1.CheckedItems.Count == 0)
                MessageBox.Show("請輸入 教室 與 時間", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (comboBox1.SelectedItem == null && comboBox2.SelectedItem != null && checkedListBox1.CheckedItems.Count == 0)
                MessageBox.Show("請輸入 課程 與 時間", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (comboBox1.SelectedItem == null && comboBox2.SelectedItem == null && checkedListBox1.CheckedItems.Count != 0)
                MessageBox.Show("請輸入 課程 與 教室", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (comboBox1.SelectedItem == null && comboBox2.SelectedItem != null && checkedListBox1.CheckedItems.Count != 0)
                MessageBox.Show("請輸入 課程", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (comboBox1.SelectedItem != null && comboBox2.SelectedItem == null && checkedListBox1.CheckedItems.Count != 0)
                MessageBox.Show("請輸入 教室", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (comboBox1.SelectedItem != null && comboBox2.SelectedItem != null && checkedListBox1.CheckedItems.Count == 0)
                MessageBox.Show("請輸入 時間", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
        //輸入課堂人數
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            string clas = getclass + " where 人數 >= " + textBox1.Text;
            DBclass(clas);
        }

        //分析
        private void button2_Click(object sender, EventArgs e)
        {
            Colorinit();
            string nowlevel = getlevel + "\"" + comboBox1.SelectedItem + "\"";
            try
            {
                string time = gettime + "\"" + label11.Text + "\"";
                DBtime(time);
                for (int column = 0; column < 5; column++)
                {
                    for (int row = 0; row < 10; row++)
                    {
                        dataGridView5[column, row].Style.BackColor = Color.Lime;
                        if (dataGridView1[column, row].Value != null
                        || dataGridView2[column, row].Value != null
                        || dataGridView3[column, row].Value != null
                        || dataGridView4[column, row].Value != null)
                        {
                            int nowlv = DBlevel(nowlevel);
                            if (dataGridView1[column, row].Value != null)
                            {
                                string[] temp = dataGridView1[column, row].Value.ToString().Split();

                                string setlevel = getlevel + "\"" + temp[0] + "\"";
                                //Level比對
                                int setlv = DBlevel(setlevel);
                                nowlv = Level_Compare(nowlv, setlv);
                                switch (nowlv)
                                {
                                    case 2:
                                        dataGridView5[column, row].Style.BackColor = Color.RoyalBlue;
                                        break;
                                    case 3:
                                        dataGridView5[column, row].Style.BackColor = Color.Yellow;
                                        break;
                                    case 4:
                                        dataGridView5[column, row].Style.BackColor = Color.Red;
                                        break;
                                    default:
                                        dataGridView5[column, row].Style.BackColor = Color.Lime;
                                        break;
                                }
                            }
                            if (dataGridView2[column, row].Value != null)
                            {
                                string[] temp = dataGridView2[column, row].Value.ToString().Split();
                                string setlevel = getlevel + "\"" + temp[0] + "\"";
                                //Level比對
                                int setlv = DBlevel(setlevel);
                                nowlv = Level_Compare(nowlv, setlv);
                                switch (nowlv)
                                {
                                    case 2:
                                        dataGridView5[column, row].Style.BackColor = Color.RoyalBlue;
                                        break;
                                    case 3:
                                        dataGridView5[column, row].Style.BackColor = Color.Yellow;
                                        break;
                                    case 4:
                                        dataGridView5[column, row].Style.BackColor = Color.Red;
                                        break;
                                    default:
                                        dataGridView5[column, row].Style.BackColor = Color.Lime;
                                        break;
                                }
                            }
                            if (dataGridView3[column, row].Value != null)
                            {
                                string[] temp = dataGridView3[column, row].Value.ToString().Split();
                                string setlevel = getlevel + "\"" + temp[0] + "\"";
                                //Level比對
                                int setlv = DBlevel(setlevel);
                                nowlv = Level_Compare(nowlv, setlv);
                                switch (nowlv)
                                {
                                    case 2:
                                        dataGridView5[column, row].Style.BackColor = Color.RoyalBlue;
                                        break;
                                    case 3:
                                        dataGridView5[column, row].Style.BackColor = Color.Yellow;
                                        break;
                                    case 4:
                                        dataGridView5[column, row].Style.BackColor = Color.Red;
                                        break;
                                    default:
                                        dataGridView5[column, row].Style.BackColor = Color.Lime;
                                        break;
                                }
                            }
                            if (dataGridView4[column, row].Value != null)
                            {
                                string[] temp = dataGridView4[column, row].Value.ToString().Split();
                                string setlevel = getlevel + "\"" + temp[0] + "\"";
                                //Level比對
                                int setlv = DBlevel(setlevel);
                                nowlv = Level_Compare(nowlv, setlv);
                                switch (nowlv)
                                {
                                    case 2:
                                        dataGridView5[column, row].Style.BackColor = Color.RoyalBlue;
                                        break;
                                    case 3:
                                        dataGridView5[column, row].Style.BackColor = Color.Yellow;
                                        break;
                                    case 4:
                                        dataGridView5[column, row].Style.BackColor = Color.Red;
                                        break;
                                    default:
                                        dataGridView5[column, row].Style.BackColor = Color.Lime;
                                        break;
                                }
                            }
                        }
                        if (comboBox2.SelectedItem != null && dataGridView6[column, row].Value != null)
                        {
                            if (dataGridView6[column, row].Value.ToString().Contains(comboBox2.SelectedItem.ToString()))
                                dataGridView5[column, row].Value = "此教室已使用";
                            else
                                dataGridView5[column, row].Value = null;
                        }
                    }
                }
                foreach (var item in DBtime(time))
                {
                    string weekstr = item.Substring(0, 1);
                    string timestr = item.Substring(1);
                    int weeknum = int.Parse(weekstr) - 1;
                    int timenum = int.Parse(timestr) - 1;
                    dataGridView5[weeknum, timenum].Style.BackColor = Color.White;
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show("No coures has chossen. " , "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            delcourse(dataGridView1);
            delcourse(dataGridView2);
            delcourse(dataGridView3);
            delcourse(dataGridView4);

        }
        //刪除課表
        private void delcourse(DataGridView dataGrid)
        {
            string[] course = textBox2.Text.Split();
            for (int Column = 0; Column < 5; Column++)
            {
                for (int Row = 0; Row < 10; Row++)
                {
                    if (dataGrid[Column, Row].Value != null && textBox2.Text != null)
                    {
                        if (dataGrid[Column, Row].Value.ToString().Contains(textBox2.Text) == true)
                        {
                            Match match = Regex.Match(textBox2.Text, @"([0-9]+[A-Z])|[0-9]+");
                            string replaceroom = match.Value + ",";//1111,
                            string lastroom = this.dataGridView6[Column, Row].Value.ToString();//1109,1102
                            Regex r = new Regex(replaceroom);
                            lastroom = r.Replace(lastroom, "");
                            this.dataGridView6[Column, Row].Value = lastroom;
                            dataGrid[Column, Row].Value = null;
                            isinsert.Remove(course[0]);
                        }
                    }
                }
            }
        }
        //將課程名稱放進textbox
        //---------------------------------------------------------------------------
        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                textBox2.Text = dataGridView1.SelectedCells[0].Value.ToString();
            }
            catch { }
        }
        private void dataGridView2_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                textBox2.Text = dataGridView2.SelectedCells[0].Value.ToString();
            }
            catch { }
        }
        private void dataGridView3_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                textBox2.Text = dataGridView3.SelectedCells[0].Value.ToString();
            }
            catch { }
        }
        private void dataGridView4_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                textBox2.Text = dataGridView4.SelectedCells[0].Value.ToString();
            }
            catch { }
        }
        //--------------------------------------------------------------------------
        //-------test------
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern bool GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern bool WritePrivateProfileString(string section, string key, string Value, string filePath);
        //讀取ini檔
        private void Getini()
        {
            StringBuilder timetemp = new StringBuilder();
            foreach (string key in coursearray)
            {
                foreach (string clas in classarray)
                {
                    if (GetPrivateProfileString("1061", key + " " + clas, "", timetemp, 255, openFileDialog1.FileName))
                    {
                        string[] t = timetemp.ToString().Split();
                        //MessageBox.Show(timetemp.ToString());
                        string grade = "select 年級 from course where 課名 =\"" + key + "\"";
                        switch (DBgrade(grade))
                        {
                            case "一":
                                for (int i = 0; i < t.Length; i++)
                                {
                                    string roomstr = t[i];
                                    string weekstr = roomstr.Substring(0, 1);
                                    string timestr = roomstr.Substring(1);
                                    int weeknum = int.Parse(weekstr) - 1;
                                    int timenum = int.Parse(timestr) - 1;
                                    DataGridViewRowCollection rows = dataGridView1.Rows;
                                    if (this.dataGridView1[weeknum, timenum].Value == null)
                                    {
                                        this.dataGridView1[weeknum, timenum].Value = key + " " + clas;
                                        this.dataGridView6[weeknum, timenum].Value += clas + ",";
                                        isinsert.Add(key);
                                    }
                                }
                                break;
                            case "二":
                                for (int i = 0; i < t.Length; i++)
                                {
                                    string roomstr = t[i];
                                    string weekstr = roomstr.Substring(0, 1);
                                    string timestr = roomstr.Substring(1);
                                    int weeknum = int.Parse(weekstr) - 1;
                                    int timenum = int.Parse(timestr) - 1;
                                    DataGridViewRowCollection rows = dataGridView2.Rows;
                                    if (this.dataGridView2[weeknum, timenum].Value == null)
                                    {
                                        this.dataGridView2[weeknum, timenum].Value = key + " " + clas;
                                        this.dataGridView6[weeknum, timenum].Value += clas + ",";
                                        isinsert.Add(key);
                                    }
                                }
                                break;
                            case "三":
                                for (int i = 0; i < t.Length; i++)
                                {
                                    string roomstr = t[i];
                                    string weekstr = roomstr.Substring(0, 1);
                                    string timestr = roomstr.Substring(1);
                                    int weeknum = int.Parse(weekstr) - 1;
                                    int timenum = int.Parse(timestr) - 1;
                                    DataGridViewRowCollection rows = dataGridView3.Rows;
                                    if (this.dataGridView3[weeknum, timenum].Value == null)
                                    {
                                        this.dataGridView3[weeknum, timenum].Value = key + " " + clas;
                                        this.dataGridView6[weeknum, timenum].Value += clas + ",";
                                        isinsert.Add(key);
                                    }
                                }
                                break;
                            case "四":
                                for (int i = 0; i < t.Length; i++)
                                {
                                    string roomstr = t[i];
                                    string weekstr = roomstr.Substring(0, 1);
                                    string timestr = roomstr.Substring(1);
                                    int weeknum = int.Parse(weekstr) - 1;
                                    int timenum = int.Parse(timestr) - 1;
                                    DataGridViewRowCollection rows = dataGridView4.Rows;
                                    if (this.dataGridView4[weeknum, timenum].Value == null)
                                    {
                                        this.dataGridView4[weeknum, timenum].Value = key + " " + clas;
                                        this.dataGridView6[weeknum, timenum].Value += clas + ",";
                                        isinsert.Add(key);
                                    }
                                }
                                break;
                        }
                    }
                }

            }


        }
        //寫入ini檔
        private void Writeini(DataGridView gridView)
        {
            List<string> key = new List<string>();
            string value;
            for (int Column = 0; Column < 5; Column++)
            {
                for (int Row = 0; Row < 10; Row++)
                {
                    if(gridView[Column, Row].Value != null)
                    {
                        if(!key.Contains(gridView[Column, Row].Value.ToString()))
                            key.Add(gridView[Column, Row].Value.ToString());
                        if((Row + 1)==10)
                            value = (Column + 1).ToString() + (Row + 1).ToString();
                        else
                            value = (Column + 1).ToString() +"0"+ (Row + 1).ToString();
                        //value append
                        if (key.Contains(gridView[Column, Row].Value.ToString())) 
                        {
                            StringBuilder timetemp = new StringBuilder();
                            GetPrivateProfileString("1061", gridView[Column, Row].Value.ToString(),"", timetemp, 255, saveFileDialog1.FileName);
                            if ((Row + 1) == 10)
                                value = timetemp.ToString() +" "+ (Column + 1).ToString() + (Row + 1).ToString();
                            else
                                value = timetemp.ToString() +" "+ (Column + 1).ToString() + "0" + (Row + 1).ToString();
                        }
                        WritePrivateProfileString("1061", gridView[Column, Row].Value.ToString(), value, saveFileDialog1.FileName);
                    }
                }
            }
        }
        //匯入
        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "ini(*.ini)|*.ini|All files (*.*)|*.*";
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                GridViewinit(dataGridView1);
                GridViewinit(dataGridView2);
                GridViewinit(dataGridView3);
                GridViewinit(dataGridView4);
                GridViewinit(dataGridView6);
                Getini();
            }         
        }
        //匯出
        private void button5_Click(object sender, EventArgs e)
        {         
            saveFileDialog1.Filter = "ini(*.ini)|*.ini|All files (*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(saveFileDialog1.FileName))
                {
                    FileInfo fi = new FileInfo(saveFileDialog1.FileName);
                    fi.Delete();
                }
                Writeini(dataGridView1);
                Writeini(dataGridView2);
                Writeini(dataGridView3);
                Writeini(dataGridView4);
            }
            saveFileDialog1.FileName = "";
        }

        //-------test------
    }
    public class Global
    {
        public static string semester = "";
    }
}
