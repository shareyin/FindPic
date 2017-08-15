using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.OleDb;

namespace Findpic
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s_D1 = textBox1.Text;
            string s_D2 = Regex.Replace(s_D1, @"\d{2}:\d{2}:\d{2}:\d{3}", "");
            string s_Dout = Regex.Replace(s_D1, @":\d{3}", "");
            s_Dout = Regex.Replace(s_Dout, "-", "");
            s_Dout = Regex.Replace(s_Dout, ":", "");
            s_Dout = Regex.Replace(s_Dout, " ", "");
            //FindAllPic(s_D2, s_Dout);

        }
        public int coun = 0;
        public void FindAllPic(string s_Data, string MarchStr, int CarNo,string forceTime,int Count)
        {

            string s_OPath = "E:\\HWCarm\\192.168.0.168\\" +s_Data;
            DirectoryInfo did = new DirectoryInfo(s_OPath);
            FileInfo[] fissd = did.GetFiles("*.jpg", SearchOption.AllDirectories);
            //先备份出来
            foreach (FileInfo fiid in fissd)
            {
                if (Regex.IsMatch(fiid.Name, MarchStr))
                {
                    File.Copy(s_OPath + "\\" + fiid.Name, "E:\\HWCarm\\LOSTALL\\" + fiid.Name);
                }

            }
            string s_Path = "";
            //s_Path = "E:\\HWCarm\\192.168.0.168\\" + s_Data;
            s_Path = "E:\\HWCarm\\LOSTALL\\";
            DirectoryInfo di = new DirectoryInfo(s_Path);
            FileInfo[] fiss = di.GetFiles("*-s.jpg", SearchOption.AllDirectories);
            string F1Name = "E:\\HWCarm\\BAK\\";
            string F2Name = "E:\\HWCarm\\LOST\\";
            if (!Directory.Exists(F1Name))
            {
                Directory.CreateDirectory(F1Name);
            }
            if (!Directory.Exists(F2Name))
            {
                Directory.CreateDirectory(F2Name);
            }
            foreach (FileInfo fii in fiss)
            {
                if (Regex.IsMatch(fii.Name, MarchStr))
                {
                    File.Move(s_Path + "\\" + fii.Name, "E:\\HWCarm\\BAK\\" + fii.Name);
                    //WriteLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "找到车牌图片：" + fii.Name+"\r\n");
                }

            }
            FileInfo[] fis = di.GetFiles("*.jpg", SearchOption.AllDirectories);
            foreach (FileInfo fi in fis)
            {
                if (Regex.IsMatch(fi.Name, MarchStr))
                {
                    string CarNum = fi.Name.Substring(fi.Name.Length - 12);
                    CarNum = Regex.Replace(CarNum, @".jpg", "");
                    string CarColor = CarNum.Substring(0, 1);
                    CarNum = CarNum.Remove(0, 1);
                    //MessageBox.Show(CarNum);
                    WriteLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "找到抓拍图片：" + fi.Name + "\r\n");
                    coun = dataGridViewout.Rows.Add();
                    CarNo++;
                    //MessageBox.Show(fi.Name);
                    if (CarColor == "蓝")
                    {
                        dataGridViewout.Rows[coun].Cells[0].Value = "客1";
                        dataGridViewout.Rows[coun].Cells[5].Value = 4911;
                        dataGridViewout.Rows[coun].Cells[6].Value = 1512;
                    }
                    else
                    {
                        dataGridViewout.Rows[coun].Cells[0].Value = "客3";
                        dataGridViewout.Rows[coun].Cells[5].Value = 9088;
                        dataGridViewout.Rows[coun].Cells[6].Value = 3739;
                    }
                    dataGridViewout.Rows[coun].Cells[1].Value = CarColor;
                    dataGridViewout.Rows[coun].Cells[2].Value = CarNum;
                    dataGridViewout.Rows[coun].Cells[3].Value = forceTime;
                    dataGridViewout.Rows[coun].Cells[4].Value = "E:\\HWCarm\\192.168.0.168\\" + s_Data+"\\"+ fi.Name;
                    
                    dataGridViewout.Rows[coun].Cells[7].Value = "改";
                    dataGridViewout.Rows[coun].Cells[8].Value = CarNo;
                    dataGridViewout.Rows[coun].Cells[9].Value = Count;
                    File.Copy(s_Path + "\\" + fi.Name, "E:\\HWCarm\\LOST\\" + fi.Name);
                    WriteLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "复制抓拍图片：" + "E:\\HWCarm\\LOST\\" + fi.Name + "\r\n");
                }

            }
        }
        //获取Excel
        public static DataSet GetExcelData(string str)
        {
            string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + str + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'";
            OleDbConnection myConn = new OleDbConnection(strCon);
            string strCom = " SELECT * FROM [车道1$]";
            myConn.Open();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
            DataSet myDataSet = new DataSet();
            myCommand.Fill(myDataSet, "[车道1$]");
            myConn.Close();
            return myDataSet;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog filedialog = new OpenFileDialog();
            string FileName = "";

            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                FileName = filedialog.FileName;
                //((DataGridViewTextBoxColumn)dataGridView1.Columns[5]).MaxInputLength=1024;
                dataGridView1.DataSource = GetExcelData(FileName);
                dataGridView1.DataMember = "[车道1$]";

                for (int count = 0; (count <= (dataGridView1.Rows.Count - 1)); count++)
                {
                    dataGridView1.Rows[count].HeaderCell.Value = (count + 1).ToString();
                }
            }
        }
        public string GetRandFor3()
        {
            Random ro = new Random();
            int iR;
            int iUp = 999;
            int iDowm = 100;
            iR = ro.Next(iDowm, iUp);
            return iR.ToString();
        }
        private void btnFind_Click(object sender, EventArgs e)
        {
            btnFind.Enabled = false;
            string CarNo = "";
            string FTimestring = "";
            string ShouldNum = "";
            
            try
            {
                for (int i = 0; i < dataGridView1.RowCount-1; i++)
                {
                    int ShijiCount = 0;
                    FTimestring = dataGridView1.Rows[i].Cells[6].Value.ToString();
                    CarNo = (this.dataGridView1.Rows[i].Cells[1].Value).ToString();
                    ShouldNum = (this.dataGridView1.Rows[i].Cells[5].Value).ToString();
                    if (FTimestring != "")
                    {
                        string[] sArry = Regex.Split(FTimestring, "    ", RegexOptions.IgnoreCase);
                        int Num = Convert.ToInt32(CarNo);
                        foreach (string ss in sArry)
                        {
                            if (ss != null || ss != "")
                            {
                                string s_D1 = ss;
                                int i_Count = 0;
                                string s_D2 = Regex.Replace(s_D1, @"\d{2}:\d{2}:\d{2}:\d{3}", "");
                                if (s_D1.Length < 3)
                                { }
                                else
                                {
                                    string s_Last3 = s_D1.Substring(s_D1.Length - 3);
                                    i_Count = Convert.ToInt16(s_Last3);
                                }
                                
                                

                                string s_Dout = Regex.Replace(s_D1, @":\d{3}", "");
                                string s_Rtime = Regex.Replace(s_D1, @":\d{3}", ":"+GetRandFor3());
                                s_Dout = Regex.Replace(s_Dout, "-", "");
                                s_Dout = Regex.Replace(s_Dout, ":", "");
                                s_Dout = Regex.Replace(s_Dout, " ", "");
                                if (i_Count > 0 && i_Count < 10)
                                {
                                    s_Dout = (Convert.ToInt64(s_Dout) - 1).ToString();
                                }
                                if (s_D2 == "" || s_Dout == "")
                                { 
                                }

                                else
                                {
                                    ShijiCount = sArry.Length-1;
                                    WriteLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "找到抓拍文件夹：" + s_D2 + "开始CarNo：" + CarNo + "开始CarNo：" + CarNo + "\r\n");
                                    FindAllPic(s_D2, s_Dout, Num,s_Rtime,ShijiCount);
                                    Num++;
                                }
                                
                                
                            }

                        }

                    }

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "查找异常：" + ex.ToString() + "\r\n");
                //btnFind.Enabled = true;
            }
            btnFind.Enabled = true;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = DateTime.Now.ToString("yyyyMMdd HHmm");
            string PathExcel = "";

            saveFileDialog1.Filter = "Excel files(*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                PathExcel = saveFileDialog1.FileName;
                ExcelCreate excre = new ExcelCreate();
                DataTable dt = new DataTable();
                dt.Columns.Add("JCarType");
                dt.Columns.Add("CarColour");
                dt.Columns.Add("PlateNo");
                dt.Columns.Add("forcetime");
                dt.Columns.Add("imagepath");
                dt.Columns.Add("carlength");
                dt.Columns.Add("carheight");
                dt.Columns.Add("update_flag");
                dt.Columns.Add("CarNo");
                dt.Columns.Add("InPutCount");

                for (int i = 0; i < dataGridViewout.Rows.Count; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr["JCarType"] = (dataGridViewout.Rows[i].Cells[0].Value);
                    dr["CarColour"] = (dataGridViewout.Rows[i].Cells[1].Value);
                    dr["PlateNo"] = (dataGridViewout.Rows[i].Cells[2].Value);
                    dr["forcetime"] = (dataGridViewout.Rows[i].Cells[3].Value);
                    dr["imagepath"] = (dataGridViewout.Rows[i].Cells[4].Value);
                    dr["carlength"] = (dataGridViewout.Rows[i].Cells[5].Value);
                    dr["carheight"] = (dataGridViewout.Rows[i].Cells[6].Value);
                    dr["update_flag"] = (dataGridViewout.Rows[i].Cells[7].Value);
                    dr["CarNo"] = (dataGridViewout.Rows[i].Cells[8].Value);
                    dr["InPutCount"] = (dataGridViewout.Rows[i].Cells[9].Value);
                    dt.Rows.Add(dr);
                }

                string dt12 = DateTime.Now.ToString("yyyyMMddhhmmss");

                excre.OutFileToDisk(dt, "数据表", PathExcel);
                MessageBox.Show("Execl导出成功，路劲为：" + PathExcel);
            }
        }

        //写日志
        public static void WriteLog(string Msg)
        {
            // string sDirPath = ".\\log\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\";
            string sDirPath = ".\\log\\";
            CreateDir(sDirPath);
            string sFilePath = sDirPath + DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            try
            {
                using (StreamWriter sw = File.AppendText(sFilePath))
                {
                    //sw.WriteLine(Msg);
                    sw.Write(Msg);
                    sw.Flush();
                    sw.Dispose();
                }
            }
            catch { }
        }
        public static void CreateDir(string dirPath)
        {
            if (!Directory.Exists(dirPath))
            {
                try
                {
                    Directory.CreateDirectory(dirPath);
                }
                catch { }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
