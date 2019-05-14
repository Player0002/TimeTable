using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;
namespace 시간표
{
    public partial class Form1 : MaterialForm
    {
        private bool checkeds;
        object[][] args = new object[5][];
        DataGridViewCell cells;
        string curremtData;
        Dictionary<string, string> dicts = new Dictionary<string, string>();
        private void Initia()
        {
            label5.Hide();
            richTextBox1.Hide();
            Save.Hide();
            MaximumSize = new Size(680, 310);
            Size = MaximumSize;
            MinimumSize = Size;
            label3.Text = "";
            dataGridView1.Rows.Clear();
            var  datac = new ExcelData();
            var args1 = datac.ReadExcelData(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\시간표.xlsx");
            for (var i = 0; i < 5; i++)
            {
                args[i] = new object[8];
            }
            for (var i = 0; i < 5; i++)
            {
                for (var j = 0; j < 8; j++)
                {
                    args[i][j] = args1.Count > 0 ? args1.Dequeue() : " ";
                }
            }
            if (!checkeds)
            {
                dataGridView1.ColumnCount = 8;
                dataGridView1.Columns[0].Name = " ";
                dataGridView1.Columns[0].Width = 656 / 8 - 22;
                for (var i = 1; i <= 7; i++)
                {
                    dataGridView1.Columns[i].Name = i + " 교시";
                    dataGridView1.Columns[i].Width = 656 / 8;
                }
                dataGridView1.AutoResizeColumns();
            }
            for (var i = 0; i < 5; i++)
            {
                dataGridView1.Rows.Add(args[i] != null ? args[i] : new object[8]);
            }
            for (var i = 0; i < 5; i++)
            {
                for (var j = 0; j < 8; j++)
                {
                    var cell = dataGridView1.Rows[i].Cells[j];
                    try
                    {
                        cell.ToolTipText = cell.Value.ToString() == "━▷" ? GetTeacher(dataGridView1.Rows[i].Cells[j - 1].Value.ToString()) != "NULL" ? "\n선생님 : " + GetTeacher(dataGridView1.Rows[i].Cells[j - 1].Value.ToString()) + "\n" : "None" : GetTeacher(cell.Value.ToString()) != "NULL" ? "\n선생님 : " + GetTeacher(cell.Value.ToString()) + "\n" : "None";
                    }
                    catch (NullReferenceException)
                    {
                    }
                }
            }

            checkeds = true;
        }
        public Form1()
        {
            InitializeComponent();
            timer1.Interval = 1000;
            timer1.Enabled = true;
            timer1.Tick += Timer1_Tick;
            dicts.Add("수학", "나요섭");
            dicts.Add("영어", "김정숙");
            dicts.Add("영어(원)", "Mr.Anderson");
            dicts.Add("과학", "이순동");
            dicts.Add("자율", "신희송");
            dicts.Add("컴일1", "장호태");
            dicts.Add("컴일2", "권태현");
            dicts.Add("C프1", "이혜인");
            dicts.Add("C프2", "신희송");
            dicts.Add("체육", "안영세");
            dicts.Add("국어", "김보연");
            dicts.Add("미술", "이진희");
            dicts.Add("동아리", "신희송");
            dicts.Add("봉사", "신희송");
            dicts.Add("컴그", "신희송");
            dicts.Add("진로", "박정렬");
            dicts.Add("사회", "김상희");
            Initia();
            var ms = MaterialSkinManager.Instance;
            ms.AddFormToManage(this);
            ms.Theme = MaterialSkinManager.Themes.LIGHT;
            ms.ColorScheme = new ColorScheme(Primary.Amber600, Primary.Amber300, Primary.LightGreen300, Accent.LightGreen200, TextShade.WHITE);
        }
        private void TimerEvent()
        {
            dataGridView1.ClearSelection();
            label1.Text = DateTime.Now.ToLongDateString().Split(' ')[3] + " " + getClassTime(DateTime.Now);
            label2.Text = DateTime.Now + " ";
            var d = Convert.ToInt32(isContains(DateTime.Now) ? Convert.ToInt32(getClassTime(DateTime.Now).Split('교')[0]) : -1);
            var rowindex = GetIndex(DateTime.Now) - 1 > 0 ? GetIndex(DateTime.Now) - 1 : 0;
            var celindex = d > 0 && d < 8 ? d : 0;
            if(GetIndex(DateTime.Now) > 0) { 
                if (d == 8) dataGridView1.Rows[rowindex].Cells[celindex - 1].Style.BackColor = Color.White;
                if (GetIndex(DateTime.Now) >= 0) dataGridView1.Rows[GetIndex(DateTime.Now) - 1].Cells[0].Style.BackColor = Color.MediumSlateBlue;
            }
            for (var i = 0; i < 5; i++) for (var j = 0; j < 8; j++) dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.White;
            if (cells != null) cells.Style.BackColor = Color.LightGray;
            if (GetIndex(DateTime.Now) <= 0) return;
            if (getClassTime(DateTime.Now).Contains("점심시간"))
            {
                dataGridView1.Rows[GetIndex(DateTime.Now) - 1].Cells[4].Style.BackColor = Color.White;
                dataGridView1.Rows[GetIndex(DateTime.Now) - 1].Cells[5].Style.BackColor = Color.Pink;
            }

            if (GetIndex(DateTime.Now) <= 0 || d <= 0) return;
            if (celindex >= 1 && dataGridView1.Rows[rowindex].Cells[celindex - 1] != cells)
            {
                dataGridView1.Rows[rowindex].Cells[celindex - 1].Style.BackColor = Color.White;
            }
            dataGridView1.Rows[rowindex < 0 ? 0 : rowindex].Cells[celindex < 0 ? 0 : celindex].Style.BackColor = Color.Aqua;
            if (celindex < 7 && getClassTime(DateTime.Now).Contains("쉬는시간")) dataGridView1.Rows[rowindex].Cells[celindex + 1].Style.BackColor = Color.Pink;
        }
        private void Timer1_Tick(object sender, EventArgs e)
        {
            TimerEvent();
        }
        bool isContains(DateTime time)
        {
            var data = getClassTime(DateTime.Now);
            var data1 = data.Contains("1") || data.Contains("2") || data.Contains("3") || data.Contains("4") || data.Contains("5") || data.Contains("6") || data.Contains("7");
            var data2 = data1 && !data.Contains("자율");
            return data2;
        }
        int GetIndex(DateTime time)
        {
            var data = time.ToString("ddd");
            if (data.Contains("월")) return 1;
            if (data.Contains("화")) return 2;
            if (data.Contains("수")) return 3;
            if (data.Contains("목")) return 4;
            return data.Contains("금") ? 5 : 0;
        }
        string getClassTime(DateTime time)
        {
            var h = Convert.ToInt32(time.ToString("HH"));
            var m = Convert.ToInt32(time.ToString("mm"));
            switch (h)
            {
                case 6 when (m >= 20 && m <= 30):
                    return "기상및 이동시간";
                case 6 when m > 30:
                case 7 when m < 21:
                    return "아침 운동시간";
                case 7 when m > 20:
                case 8 when m <= 20:
                    return "아침식사후 기숙사 나가기";
                case 8 when m > 20 && m <= 30:
                    return "학교 등교시간";
                case 8 when m > 30 && m < 40:
                    return "아침 조회";
                case 8 when m >= 40:
                case 9 when m < 30:
                    return "1교시";
                case 9 when m >= 30 && m < 40:
                    return "1교시 쉬는시간";
                case 9 when m >= 40:
                case 10 when m < 30:
                    return "2교시";
                case 10 when m >= 30 && m < 40:
                    return "2교시 쉬는시간";
                case 10 when m >= 40:
                case 11 when m < 30:
                    return "3교시";
                case 11 when m >= 30 && m < 40:
                    return "3교시 쉬는시간";
                case 11 when m >= 40:
                case 12 when m < 30:
                    return "4교시";
                case 12 when m >= 30 && m < 40:
                    return "4교시 쉬는시간 + 점심";
                case 12 when m >= 40:
                case 13 when m < 20:
                    return "점심시간";
                case 13 when m >= 20:
                case 14 when m < 10:
                    return "5교시";
                case 14 when m >= 10 && m < 20:
                    return "5교시 쉬는시간";
                case 14 when m >= 20:
                case 15 when m < 10:
                    return "6교시";
                case 15 when m >= 10 && m < 20:
                    return "6교시 쉬는시간";
                case 15 when m >= 20:
                case 16 when m < 10:
                    return "7교시";
                case 16 when m >= 10 && m < 30:
                    return "7교시 쉬는시간";
                case 16 when m >= 30:
                case 17 when m < 20:
                    return "8교시";
                case 17 when m >= 20 && m < 30:
                    return "8교시 쉬는시간";
                case 17 when m >= 30:
                case 18 when m < 20:
                    return "9교시";
                case 18 when m >= 20:
                case 19 when m < 10:
                    return "저녁시간";
                case 19 when m >= 10 && m <= 59:
                    return "자율 1교시";
                case 20 when m >= 0 && m < 10:
                    return "자율 1 쉬는시간";
                case 20 when m >= 10 && m <= 59:
                    return "자율 2교시";
                case 21 when m >= 0 && m < 20:
                    return "기숙사 이동";
                case 21 when m >= 20 && m < 40:
                    return "개인시간";
                case 21 when m >= 40 && m < 50:
                    return "인원확인 (점호)";
                case 21 when m >= 50:
                case 22 when m < 55:
                    return "심야자습";
                case 22 when m >= 55 && m <= 59:
                    return "기숙사 이동";
                default:
                    return "취침시간";
            }
        }
        private string GetTeacher(string s)
        {

            foreach (var ss in dicts.Keys)
            {
                if (ss == s) return dicts[s];
            }
            return "NULL";
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: 이 코드는 데이터를 'dataBaseDataSet.Sheet1' 테이블에 로드합니다. 필요 시 이 코드를 이동하거나 제거할 수 있습니다.

        }
        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView1.Rows.Clear();
            for (var i = 0; i < 5; i++)
            {
                dataGridView1.Rows.Add(args[i] != null ? args[i] : new object[8]);
            }
            for (var i = 0; i < 5; i++)
            {
                for (var j = 0; j < 8; j++)
                {
                    var cell = dataGridView1.Rows[i].Cells[j];
                    try
                    {
                        cell.ToolTipText = cell.Value.ToString() == "━▷" ? GetTeacher(dataGridView1.Rows[i].Cells[j - 1].Value.ToString()) != "NULL" ? "\n선생님 : " + GetTeacher(dataGridView1.Rows[i].Cells[j - 1].Value.ToString()) + "\n" : "None" : GetTeacher(cell.Value.ToString()) != "NULL" ? "\n선생님 : " + GetTeacher(cell.Value.ToString()) + "\n" : "None";
                    }
                    catch (NullReferenceException)
                    {
                    }
                }
            }
            TimerEvent();
            cells.Style.BackColor = Color.LightGray;
        }
        private void Reload_Click(object sender, EventArgs e)
        {
            Initia();
        }

        private string GetCurrentData(int n)
        {
            switch (n)
            {
                case(0):
                    return "월";
                case (1):
                    return "화";
                case (2):
                    return "수";
                case (3):
                    return "목";
                case (4):
                    return "금";
                default: return "Unable";
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 1 && e.ColumnIndex >= 1)
            {
                cells = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                var datas = (cells.Value.ToString().Contains("━▷")
                    ? dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value
                    : cells.Value).ToString();
                if (GetTeacher(datas) != "NULL")
                {
                    
                    label3.Text = GetCurrentData(e.RowIndex) + " " + datas;
                    label4.Text = GetTeacher(datas);
                    curremtData = GetCurrentData(e.RowIndex) + " - " + datas;
                    isInData(e.RowIndex, datas);
                    TimerEvent();
                }
            }
        }

        private void isInData(int index, string datas)
        {

            richTextBox1.Clear();
            if (curremtData == null || !curremtData.Contains("-")) return;
            SetData(curremtData);
            Console.WriteLine(curremtData);
            this.MaximumSize = new Size(680 + 280, this.Height);
            this.Size = this.MaximumSize;
            MinimumSize = Size;
            var TextBox = richTextBox1;
            TextBox.Show();
            Save.Show();
            label5.Show();
            button1.Show();
            Loads();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            richTextBox1.Hide();
            Save.Hide();
            label5.Hide();
            label5.Hide();
            button1.Hide();
            this.MaximumSize = new Size(680, this.Height);
            this.Size = this.MaximumSize;
            MinimumSize = Size;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }
        public void SetData(string d)
        {
            label5.Text = d;
        }

        private void Save_Click(object sender, EventArgs e)
        {
            var today = label5.Text.Split('-')[0].Replace(" ", "");
            var classs = label5.Text.Split('-')[1].Replace(" ", "");
            var datas = new FileData();
            datas.Write(today, classs, richTextBox1.Text.Split('\n').Select(s => s.Replace("\\today\\", DateTime.Now.ToString())).ToArray());
            richTextBox1.Clear();
            Loads();
        }
        private void Loads()
        {
            var data = new FileData();
            var today = label5.Text.Split('-')[0].Replace(" ", "");
            var classs = label5.Text.Split('-')[1].Replace(" ", "");

            if (data.ReadLine(today, classs) == null) return;
            foreach (var current in data.ReadLine(today, classs))
            {
                richTextBox1.AppendText(current + "\n");
            }
        }
        
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
    class ExcelData
    {
        public Queue<object> ReadExcelData(string path)
        {
            if (!File.Exists(path))
            {
                Debug.Write("Error Can not find TimeTable");
                System.Environment.Exit(-1);
            }
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;
            var objs = new Queue<object>();
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                wb = excelApp.Workbooks.Open(path);
                ws = wb.Worksheets.Item[1] as Microsoft.Office.Interop.Excel.Worksheet;
                var rng = ws.UsedRange;
                object[,] data = ws.Range["A1", "H5"].Value;
                for (var r = 1; r <= data.GetLength(0); r++)
                {
                    for (var c = 1; c <= data.GetLength(1); c++)
                    {
                        if (data[r, c] == null) continue;
                        if (data[r, c] != null) objs.Enqueue(data[r, c]);
                    }
                }
                wb.Close(true);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw new Exception("Error : " + ex);
            }
            return objs;
        }
    }
}
