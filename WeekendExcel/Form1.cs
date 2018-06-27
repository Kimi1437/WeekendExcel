using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading;

namespace WeekendExcel
{
    public partial class Form1 : Form
    {
        public object Messagebox { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 程序加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            String constr = "data source =192.168.80.18; initial catalog = Top; user id = sa; password =sa";  //连接数据库服务器
            SqlConnection conn = new SqlConnection(constr);//SQL连接类的实例化
            conn.Open();//打开数据库
            MessageBox.Show("小伙子 我看你骨骼精奇，你居然已经连上了18数据库，连接成功");//弹出窗口，用于测试数据库连接是否成功。
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 导出人员登录统计文件按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string a = "C:" + "\\人员登录统计.xlsx";
            ExportExcels(a, dataGridView1);
        }
        /// <summary>
        /// 导出部门登录统计文件按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            string a = "C:" + "\\部门登录统计.xlsx";
            ExportExcels(a, dataGridView2);
        }

        /// <summary>
        /// 人员查询按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            Progress fm = new Progress(0, 100);
            fm.Show(this);//设置父窗体
            fm.setPos(1, "开始查询");//设置进度条位置
            Thread.Sleep(100);//睡眠时间为100
            DataTable resultData = new DataTable();//最终的表格 
            DataTable dataData = new DataTable();//临时的表格数据
            string sql = "";
            string  dateBegin = Convert.ToString(dateTimePicker1.Text).Replace("年","-").Replace("月","-").Replace("日","");
            string dateEnd = Convert.ToString(dateTimePicker2.Text).Replace("年", "-").Replace("月", "-").Replace("日", "");
          
            DateTime dateBeginTime = Convert.ToDateTime(dateBegin);
            DateTime dateEndTime = Convert.ToDateTime(dateEnd);
            fm.setPos(10, "准备查询数据");//设置进度条位置
            Thread.Sleep(100);//睡眠时间为100
            TimeSpan sp = dateEndTime.Subtract(dateBeginTime);
            int number = 1;
            for (int j = 0; j < sp.Days; j++)
            {              
                sql = $@"select count(1) as loginnumber,  useraccount,usercnname,locdepartmentcnname 
                from (
                select  modifytime, substring(modifyuser, charindex('/', modifyuser) + 1, 100) as Loginuser
                    from BPMAPP_LOG where type = 0
                    )a inner join WECHATUSER b on a.Loginuser = b.useraccount
                    where a.modifytime >= '" + dateBeginTime.ToString("yyyy-MM-dd") + "' and a.modifytime < '" + dateBeginTime.AddDays(1).ToString("yyyy-MM-dd") + "' and a.Loginuser <>'zhekans'group by useraccount,usercnname,locdepartmentcnname";//执行的查询语句
                fm.setPos(30+j, "开始拼了命的查询数据");//设置进度条位置
                Thread.Sleep(100);//睡眠时间为100
                dataData = ConSQL(sql, dateBeginTime.ToString("yyyy-MM-dd"));
                if (number==1)
                {
                    resultData = dataData.Clone();
                }
                object[] obj = new object[dataData.Columns.Count];
                for (int i = 0; i < dataData.Rows.Count; i++)
                {
                    dataData.Rows[i].ItemArray.CopyTo(obj, 0);
                    resultData.Rows.Add(obj);
                }
                dateBeginTime = dateBeginTime.AddDays(1);
                number++;
                fm.setPos(60 + j, "开始拼了命的查询数据");//设置进度条位置
                Thread.Sleep(100);//睡眠时间为100
            }
            #region 更换列的顺序以及修改列名
            fm.setPos(90 , "组装数据中…………");//设置进度条位置
            Thread.Sleep(100);//睡眠时间为100
            resultData.Columns["date"].SetOrdinal(0);
            resultData.Columns["date"].ColumnName = "日期";
            resultData.Columns["loginnumber"].SetOrdinal(1);
            resultData.Columns["loginnumber"].ColumnName = "登录次数";
            resultData.Columns["useraccount"].SetOrdinal(2);
            resultData.Columns["useraccount"].ColumnName = "用户账号";
            resultData.Columns["usercnname"].SetOrdinal(3);
            resultData.Columns["usercnname"].ColumnName = "用户名";
            resultData.Columns["locdepartmentcnname"].SetOrdinal(4);
            resultData.Columns["locdepartmentcnname"].ColumnName = "登录部门";
            #endregion

            dataGridView1.DataSource = resultData;
            fm.Close();//关闭进度条窗体
        }

        /// <summary>
        /// 部门登录查询
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            Progress fm = new Progress(0, 100);
            fm.Show(this);//设置父窗体
            fm.setPos(1, "开始查询");//设置进度条位置
            Thread.Sleep(100);//睡眠时间为100
            DataTable resultData = new DataTable();//最终的表格 
            DataTable dataData = new DataTable();//临时的表格数据
            string sql = "";
            string dateBegin = Convert.ToString(dateTimePicker1.Text).Replace("年", "-").Replace("月", "-").Replace("日", "");
            string dateEnd = Convert.ToString(dateTimePicker2.Text).Replace("年", "-").Replace("月", "-").Replace("日", "");

            DateTime dateBeginTime = Convert.ToDateTime(dateBegin);
            DateTime dateEndTime = Convert.ToDateTime(dateEnd);
            fm.setPos(10, "准备查询数据");//设置进度条位置
            Thread.Sleep(100);//睡眠时间为100
            TimeSpan sp = dateEndTime.Subtract(dateBeginTime);
            int number = 1;
            for (int j = 0; j < sp.Days; j++)
            {
                sql = $@"select count(1)as loginnumber,  locdepartmentcnname from 
                        (
                        select  modifytime, substring(modifyuser, charindex('/',modifyuser)+1,100)as Loginuser
                          from BPMAPP_LOG where type=0 
                        )a  inner join WECHATUSER b on a.Loginuser=b.useraccount
                        where a.modifytime>='" + dateBeginTime.ToString("yyyy-MM-dd") + "' and a.modifytime<'" + dateBeginTime.AddDays(1).ToString("yyyy-MM-dd") + "'and  a.Loginuser<>'zhekans' group by locdepartmentcnname";//执行的查询语句
                fm.setPos(30 + j, "开始拼了命的查询数据");//设置进度条位置
                Thread.Sleep(100);//睡眠时间为100
                dataData = ConSQL(sql, dateBeginTime.ToString("yyyy-MM-dd"));
                if (number == 1)
                {
                    resultData = dataData.Clone();
                }
                object[] obj = new object[dataData.Columns.Count];
                for (int i = 0; i < dataData.Rows.Count; i++)
                {
                    dataData.Rows[i].ItemArray.CopyTo(obj, 0);
                    resultData.Rows.Add(obj);
                }
                dateBeginTime = dateBeginTime.AddDays(1);
                number++;
                fm.setPos(60 + j, "开始拼了命的查询数据");//设置进度条位置
                Thread.Sleep(100);//睡眠时间为100
            }
            #region 更换列的顺序以及修改列名
            fm.setPos(90, "组装数据中…………");//设置进度条位置
            Thread.Sleep(100);//睡眠时间为100
            resultData.Columns["date"].SetOrdinal(0);
            resultData.Columns["date"].ColumnName = "日期";
            resultData.Columns["loginnumber"].SetOrdinal(1);
            resultData.Columns["loginnumber"].ColumnName = "登录次数";
            resultData.Columns["locdepartmentcnname"].SetOrdinal(2);
            resultData.Columns["locdepartmentcnname"].ColumnName = "登录部门";
            #endregion
            dataGridView2.DataSource = resultData;
            fm.Close();//关闭进度条窗体
        }
        /// <summary>
        /// 连接数据库
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        public DataTable ConSQL(string sql,string date)
        {
            String constr = "data source =192.168.80.18; initial catalog = Top; user id = sa; password =sa";  //连接数据库服务器
            SqlConnection conn = new SqlConnection(constr);//SQL连接类的实例化
            conn.Open();//打开数据库
            SqlDataAdapter find = new SqlDataAdapter(sql, conn);// 创建DataAdapter数据适配器实例，SqlDataAdapter作用是 DataSet和 SQL Server之间的桥接器，用于检索和保存数据
            DataSet save = new DataSet(); //创建DataSet实例
            find.Fill(save, "wu");//  使用DataAdapter的Fill方法(填充)，调用SELECT命令 fill(对象名，"自定义虚拟表名")  
            save.Tables[0].Columns.Add("date");
            for (int i = 0; i < save.Tables[0].Rows.Count; i++)
            {
                save.Tables[0].Rows[i]["date"] = date;
            }
            //dataGridView1.DataSource = save.Tables[0];// 向DataGridView1中填充数据
            conn.Close();//关闭数据库
            return save.Tables[0];
        }
        /// <summary>
        /// 导出Excel文件
        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="myDGV">控件DataGridView</param>
        private void ExportExcels(string fileName, DataGridView myDGV)
        {
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = fileName;
            saveDialog.ShowDialog();
            saveFileName = saveDialog.FileName;
            if (saveFileName.IndexOf(":") < 0) return; //被点了取消
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("无法创建Excel对象，可能您的机子未安装Excel");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1
                                                                                                                                  //写入标题
            for (int i = 0; i < myDGV.ColumnCount; i++)
            {
                worksheet.Cells[1, i + 1] = myDGV.Columns[i].HeaderText;
            }
            //写入数值
            for (int r = 0; r < myDGV.Rows.Count; r++)
            {
                for (int i = 0; i < myDGV.ColumnCount; i++)
                {
                    worksheet.Cells[r + 2, i + 1] = myDGV.Rows[r].Cells[i].Value;
                }
                System.Windows.Forms.Application.DoEvents();
            }
            worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
            if (saveFileName != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                }
            }
            xlApp.Quit();
            GC.Collect();//强行销毁
            MessageBox.Show("文件： " + fileName + ".xlsx 保存成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        /// <summary>
        /// 跳转到操作统计
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            LogonStatistics form = new LogonStatistics();
            form.Show();
        }

        private delegate void SetPos(int ipos, string vinfo);//代理
    }
}
