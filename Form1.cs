using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
//必要性更新
//内容显示还能优化，更换为别的，输出一大块信息
//没有打开文件前不能按搜索√
//自动杀进程，不需要按关闭
//显示下一个

//花里胡哨更新
//显示还有多少单元格没搜索
//
namespace SearchOfExcel
{
    public partial class Form1 : Form
    {
        //全局变量
        Excel.Application excel;
        Excel.Workbook xBook;
        char Zimu;
        int Label_Row;
        int find;
        string Search;

        public Form1()
        {
            InitializeComponent();
        }
        //初始化部分
        private void Form1_Load(object sender, EventArgs e)
        {
            excel = new Excel.Application();
            Start.Enabled = false;
        }
        //选择文件
        private void button1_Click(object sender, EventArgs e)
        {
            Zimu = Convert.ToChar('A' - 1);
            Label_Row = 1;
            GetPath();
            Start.Enabled = true;
        }
        //搜索按键
        private void Start_Click(object sender, EventArgs e)
        {
            if (Zimu != labelBox.Text[0] || Search != SearchWord.Text)
            {
                Zimu = labelBox.Text[0];
                Search = SearchWord.Text;
                find = 0;
            }
            Searching(excel, xBook, Zimu, Label_Row, Search);
        }
        //关闭按键
        private void end_Click_1(object sender, EventArgs e)
        {
            EndOfExcel(excel, xBook);
            this.Close();
        }


        //各种子函数
        //选择文件路径
        private void GetPath()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "Excel文件(*.xlsx*)|*.xlsx*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                try
                {
                    xBook = excel.Workbooks._Open(@file);
                    FILENAME.Text = file;
                    GetLabel(excel, Zimu, ref Label_Row);
                }
                catch
                {
                    MessageBox.Show("该文件不是excel文件\n请重新选择" ,"警告");
                }
            }
        }
        //获取标签
        private void GetLabel(Excel.Application excel, char Zimu, ref int number)
        {
            int end = 0;
            labelBox.Items.Clear();
            while (end <= 3)
            {
                Zimu++;
                string position = Zimu + Convert.ToString(number);
                Excel.Range rng1 = excel.get_Range(position, Type.Missing);
                //单元格为空
                if (rng1.Value2 == null)
                {
                    end++;
                    if (end == 2 && Zimu <= 'C')//过早为null，怀疑是标题
                    {
                        //判断是否是合并单元格
                        rng1 = (Excel.Range)excel.Cells[1, "A"];
                        bool isMerge = (bool)rng1.MergeCells;
                        if (isMerge)
                        {
                            //是则判断下一行是否为标签行
                            number++;
                            Zimu = Convert.ToChar('A' - 1);
                            end = 0;
                            labelBox.Items.Clear();
                        }
                    }
                    continue;
                }
                //将标签存进box里
                labelBox.Items.Add(Zimu + ":" + rng1.Value2.ToString());
                if (Zimu == 'A')
                    labelBox.Text = (Zimu + ":" + rng1.Value2.ToString());
                end = 0;
            }
        }
        //查找
        private void Searching(Excel.Application excel, Excel.Workbook xBook, char Zimu, int Label_Row, string Search)
        {
            int number = 0;
            int end = 0;
            int i = 0;

            //从某一列中寻找
            while (true)
            {
                number++;
                string position = Zimu + Convert.ToString(number);
                Excel.Range rng1 = excel.get_Range(position, Type.Missing);
                //如果单元格为空
                if (rng1.Value2 == null)
                {
                    end++;
                    if (end == 4)
                    {
                        if (find == 0)
                            UnFound();
                        break;
                    }
                    continue;
                }
                //如果找到了一样的信息，输出
                if (Search == Convert.ToString(rng1.Value2))
                {
                    if (i == find)
                    {
                        Found(excel, number, Label_Row);
                        find++;
                        break;
                    }
                    i++;
                }
                end = 0;
            }
            //Console.WriteLine("over");
        }
        //找不到该信息（）
        static void UnFound()
        {
            MessageBox.Show("找不到该信息");
            //Console.WriteLine("找不到该信息");
        }
        //找到了，打印信息
        private void Found(Excel.Application excel, int number, int Label_Row)
        {
            int end = 0;
            char Zimu = 'A';
            Show.Clear();
            Show.AppendText("find it!in " + number);
            //Console.WriteLine("detail:");
            while (end <= 5)
            {
                string position = Zimu + Convert.ToString(number);
                string label_pos = Zimu + Convert.ToString(Label_Row);
                Excel.Range label = excel.get_Range(label_pos, Type.Missing);
                Excel.Range rng1 = excel.get_Range(position, Type.Missing);
                Zimu++;
                if (rng1.Value2 == null)
                {
                    end++;
                    continue;
                }
                else if (label.Value2 == null)
                {
                    Show.AppendText("\r\n" + " " + ":" + rng1.Value2.ToString());
                }
                else
                {
                    Show.AppendText("\r\n" + label.Value2.ToString() + ":" + rng1.Value2.ToString());
                }
                end = 0;
            }
        }
        //结束excel进程
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        static void EndOfExcel(Excel.Application excel, Excel.Workbook xBook)
        {
            //xBook.Close();
            excel.Quit();
            try
            {
                if (excel != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(excel.Hwnd), out lpdwProcessId);
                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Delete Excel Process Error:" + ex.Message);
            }
        }
    }
}
