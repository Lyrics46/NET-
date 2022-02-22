using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using MessageBox = System.Windows.Forms.MessageBox;
/// <summary>
/// .NET实验大作业  
/// 根据课程工作量统计表，统计所有老师的工作量，并形成报表。
/// 需实现功能如下：
/// 1.美观的用户UI  （待完善！）
/// 2.选择多个EXCEL   （已实现）
/// 3.可以根据教师姓名统计某位老师的工作量，并生成新的EXCEL  （已实现）
/// 4.可以统计所有老师工作量，形成新的EXCEL，并可以生成所有老师的工作总量报表  （已实现）
/// 5.实现姓名合并功能，避免有时姓名输入错误  （已实现）
/// 6.实现转存到sqlite等数据库（待完善！）
/// 
/// 其它注意事项：
/// 1.对于不同的功能所要选择的表格文件不同，请UI设计人员增加相应的提示信息，以便此项目拥有更好的用户体验。
/// 2.本项目中未实现判断选择的表格文件是否与相应功能一致的功能，欢迎完善。
/// </summary>
namespace WpfApp2
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        List<String> excelspath = new List<String>();
        //int sum = 0;//合计工作量

        public MainWindow()
        {
            InitializeComponent();
        }


        private void ButtonAddName_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.InitialDirectory = "E:\\仿真";
            //openFileDialog1.Filter = "*.xlsx|*.xls";//筛选表格文件
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            //openFileDialog1.InitialDirectory = "E:\\";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && !lstNames.Items.Contains(openFileDialog1.FileName))
            {
                lstNames.Items.Add(openFileDialog1.FileName);//显示
                excelspath.Add(openFileDialog1.FileName);//并存放在excels中
            }
        }//选择文件

        private void ButtonCountTeacher_Click(object sender, RoutedEventArgs e)
        {

            //需读取的文件是14152.xls或者141521.xlsx
            //要写的文件为teacherName.txt ，文件结构同14152.xls 
            if (!string.IsNullOrWhiteSpace(teacherName.Text))
            {
                //已选择的文件需为课程工作量统计表
                string filePath = null;
                Array excelspaths = excelspath.ToArray();
                //int countnum = excelspaths.Length;

                DataTable dt = new DataTable();
                FileStream fsRead = null;
                IWorkbook wkBook = null;
                DataColumn column = null;
                DataRow dataRow = null;
                int teacherNamerow = -1;
                for (int count = 0; count < excelspaths.Length; count++)
                {
                    filePath = excelspath[count];
                    if (filePath != null)
                    {
                        //在课程工作量统计表文件中找teacherName所在行
                        fsRead = new FileStream(filePath, FileMode.Open);//1、创建一个工作簿workBook对象
                        if (filePath.IndexOf(".xlsx") > 0)////将人员表.xls中的内容读取到fsRead中
                            wkBook = new XSSFWorkbook(fsRead);
                        else if (filePath.IndexOf(".xls") > 0)
                            wkBook = new HSSFWorkbook(fsRead);

                        for (int i = 0; i < wkBook.NumberOfSheets; i++)
                        {
                            //获取每个工作表对象
                            ISheet sheet = wkBook.GetSheetAt(i);
                            //获取每个工作表的行
                            for (int r = 0; r < sheet.LastRowNum; r++)
                            {
                                IRow currentRow = sheet.GetRow(r);
                                if (currentRow == null) continue;
                                if (teacherNamerow == r)
                                {
                                    dataRow = dt.NewRow();
                                }
                                for (int c = 0; c < currentRow.LastCellNum; c++)
                                {
                                    try
                                    {
                                        //获取每个单元格
                                        ICell cell = currentRow.GetCell(c);
                                        if (r == 0 && count == 0)//第一次读把首行读进dt里 
                                        {
                                            //本行都存入数据结构中
                                            if (cell == null)
                                            {
                                                column = new DataColumn(" ");
                                            }
                                            else
                                            {
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        column = new DataColumn(" ");
                                                        break;
                                                    case CellType.String:
                                                        column = new DataColumn(cell.StringCellValue);
                                                        break;
                                                }
                                            }
                                            dt.Columns.Add(column);
                                        }//没问题了
                                        if (cell != null && cell.CellType == CellType.String && cell.StringCellValue == teacherName.Text && teacherNamerow == -1)
                                        {
                                            //找到teacherName所在行   
                                            teacherNamerow = r;
                                            r--;
                                            break;
                                        }

                                        if (r == teacherNamerow)
                                        {
                                            //本行的存入
                                            if (cell == null)
                                            {
                                                dataRow[c] = " ";
                                            }
                                            else
                                            {
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        dataRow[c] = "";
                                                        break;
                                                    case CellType.Numeric:
                                                        short format = cell.CellStyle.DataFormat;
                                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                                            dataRow[c] = cell.DateCellValue;
                                                        else
                                                            dataRow[c] = cell.NumericCellValue;
                                                        break;
                                                    case CellType.String:
                                                        dataRow[c] = cell.StringCellValue;
                                                        break;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ;
                                    }
                                }
                                if (dataRow != null && teacherNamerow == r)
                                {
                                    dt.Rows.Add(dataRow.ItemArray);
                                }
                            }
                        }
                        fsRead.Close();
                        wkBook.Close();
                        fsRead.Dispose();
                    }//endif
                    teacherNamerow = -1;
                }
                DataTableToExcel(dt, teacherName.Text);
                dt.Dispose();
            }//endif

            //窗口功能   
            /**********************************
             * （伟大的苗皇负责部分）
             *   待完成
             *********************************/
            MessageBox.Show("已成功写入文件");
        }//生成个人文件

        private bool DataTableToExcel(DataTable dt,String filename)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new HSSFWorkbook();
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表  
                    int rowCount = dt.Rows.Count;//行数  
                    int columnCount = dt.Columns.Count;//列数  

                    //设置列头  
                    row = sheet.CreateRow(0);//excel第一行设为列头  
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,  
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据  
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                   // using (fs = File.OpenWrite(@"D:/myxls.xls"))
                   using(fs=File.OpenWrite("E://"+filename+".xls"))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据  
                        result = true;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return false;
            }


        }//存入excel

        private void ButtonCountAllTeachers_Click(object sender, RoutedEventArgs e)
        {
            /***************************************************************************************
             * 每个表格只读取名字列和总计列，
             * 读取名字列后再dt中判断是否已有该名字字段，若已有则修改它对应的总工作量值
             * 若没有则新增一行
             *****************************************************************************************/

            //已选择的文件需为课程工作量统计表
            string filePath = null;
            Array excelspaths = excelspath.ToArray();

            DataTable dt = new DataTable();
            FileStream fsRead = null;
            IWorkbook wkBook = null;
            DataColumn column = null;
            DataRow dataRow = null;
            for (int count = 0; count < excelspaths.Length; count++)
            {
                if (count == 0)//读第二个表的时候不需要再读这些东西
                {
                    column = new DataColumn("姓名");
                    dt.Columns.Add(column);
                    column = new DataColumn("工作量");
                    dt.Columns.Add(column);
                }

                filePath = excelspath[count];
                if (filePath != null)
                {
                    //若读到的姓名行中 在课程工作量统计表文件中
                    fsRead = new FileStream(filePath, FileMode.Open);//1、创建一个工作簿workBook对象
                    if (filePath.IndexOf(".xlsx") > 0)////将人员表.xls中的内容读取到fsRead中
                        wkBook = new XSSFWorkbook(fsRead);
                    else if (filePath.IndexOf(".xls") > 0)
                        wkBook = new HSSFWorkbook(fsRead);

                    for (int i = 0; i < wkBook.NumberOfSheets; i++)
                    {
                        //获取每个工作表对象
                        ISheet sheet = wkBook.GetSheetAt(i);
                        int cname=0,csum=0 ;
                        IRow FirstRow = sheet.GetRow(0);
                        for (int c = 0; c < FirstRow.LastCellNum; c++)
                        {
                            ICell cell = FirstRow.GetCell(c);
                            if (cell.CellType== CellType.String&& cell.StringCellValue=="姓名")
                            {
                                cname = c;
                            }
                            else if(cell.CellType == CellType.String && cell.StringCellValue == "总计")
                            {
                                csum = c;
                            }
                        }

                        //获取每个工作表的行
                        for (int r = 1; r <= sheet.LastRowNum; r++)
                        {
                            IRow currentRow = sheet.GetRow(r);
                            if (currentRow == null) continue;
                            //只需要判断特定列
                            ICell cell = currentRow.GetCell(cname);
                            bool flag = false;
                            int flagr;//标记bt中的哪行有对应已存在的名
                            for ( flagr = 0; flagr < dt.Rows.Count; flagr++)
                            {
                                if(dt.Rows[flagr][0].ToString()== cell.StringCellValue)
                                {
                                    flag = true;
                                    break;
                                }
                            }
                            if (!flag)
                            {
                                dataRow = dt.NewRow();//新建行
                                dataRow[0] = cell.StringCellValue;
                                cell = currentRow.GetCell(csum);
                                dataRow[1] = int.Parse(cell.NumericCellValue.ToString());//cell.NumericCellValue;
                            }
                            else
                            {
                                cell = currentRow.GetCell(csum);
                                int sum = int.Parse(dt.Rows[flagr][1].ToString());
                                sum+= int.Parse(cell.NumericCellValue.ToString());
                                dt.Rows[flagr][1]= sum;// += cell.NumericCellValue.ToString();//??????
                            }                  
                            if (dataRow != null )
                            {
                                dt.Rows.Add(dataRow.ItemArray);
                                dataRow = null;
                            }
                        }
                    }
                    fsRead.Close();
                    wkBook.Close();
                    fsRead.Dispose();
                }//endif
            }
            DataTableToExcel(dt, "汇总");
            dt.Dispose();

            //以下为界面代码
            /***界面代码****/
            MessageBox.Show("已成功写入文件");
        }//汇总

        private void ButtonMerge_Click(object sender, RoutedEventArgs e)
        {
            //已选择的文件需为需要合并的文件
            //要写的文件为rightName.xls ，文件结构同张三.xls 
            if (!string.IsNullOrWhiteSpace(rightName.Text))
            {
                string filePath = null;
                Array excelspaths = excelspath.ToArray();
                DataTable dt = new DataTable();
                FileStream fsRead = null;
                IWorkbook wkBook = null;
                DataColumn column = null;
                DataRow dataRow = null;
                for (int count = 0; count < excelspaths.Length; count++)
                {
                    filePath = excelspath[count];
                    if (filePath != null)
                    {
                        fsRead = new FileStream(filePath, FileMode.Open);//1、创建一个工作簿workBook对象
                        if (filePath.IndexOf(".xlsx") > 0)////将人员表.xls中的内容读取到fsRead中
                            wkBook = new XSSFWorkbook(fsRead);
                        else if (filePath.IndexOf(".xls") > 0)
                            wkBook = new HSSFWorkbook(fsRead);

                        for (int i = 0; i < wkBook.NumberOfSheets; i++)
                        {
                            //获取每个工作表对象
                            ISheet sheet = wkBook.GetSheetAt(i);

                            int cname = 0;
                            IRow FirstRow = sheet.GetRow(0);
                            for (int c = 0; c < FirstRow.LastCellNum; c++)
                            {
                                ICell cell = FirstRow.GetCell(c);
                                if (cell.CellType == CellType.String && cell.StringCellValue == "姓名")
                                {
                                    cname = c;
                                }
                            }

                            //获取每个工作表的行
                            for (int r = 0; r <= sheet.LastRowNum; r++)
                            {
                                IRow currentRow = sheet.GetRow(r);
                                if (currentRow == null) continue;
                                if (r != 0)
                                {
                                    dataRow = dt.NewRow();
                                }
                                for (int c = 0; c < currentRow.LastCellNum; c++)
                                {
                                    try
                                    {
                                        //获取每个单元格
                                        ICell cell = currentRow.GetCell(c);
                                        if (r == 0 && count == 0)//第一次读把首行读进dt里 
                                        {
                                            //本行都存入数据结构中
                                            if (cell == null)
                                            {
                                                column = new DataColumn(" ");
                                            }
                                            else
                                            {
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        column = new DataColumn(" ");
                                                        break;
                                                    case CellType.String:
                                                        column = new DataColumn(cell.StringCellValue);
                                                        break;
                                                }
                                            }
                                            dt.Columns.Add(column);
                                        }//没问题了

                                        if (r != 0)
                                        {
                                            //本行的存入
                                            if (cell == null)
                                            {
                                                dataRow[c] = " ";
                                            }
                                            else
                                            {
                                                switch (cell.CellType)
                                                {
                                                    case CellType.Blank:
                                                        dataRow[c] = "";
                                                        break;
                                                    case CellType.Numeric:
                                                        short format = cell.CellStyle.DataFormat;
                                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                                            dataRow[c] = cell.DateCellValue;
                                                        else
                                                            dataRow[c] = cell.NumericCellValue;
                                                        break;
                                                    case CellType.String:
                                                        if (c == cname && cell.StringCellValue != rightName.Text)
                                                            dataRow[c] = rightName.Text;
                                                        else
                                                            dataRow[c] = cell.StringCellValue;
                                                        break;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ;
                                    }
                                }
                                if (dataRow != null)
                                {
                                    dt.Rows.Add(dataRow.ItemArray);
                                    dataRow = null;
                                }
                            }
                        }
                        fsRead.Close();
                        wkBook.Close();
                        fsRead.Dispose();
                    }//endif
                }
                DataTableToExcel(dt, rightName.Text);
                dt.Dispose();
                MessageBox.Show("已成功写入文件" + rightName.Text);
            }//endif

        }//合并

        private void ButtonSaveSQL_Click(object sender, RoutedEventArgs e)//存入数据库
        {
            /**************************************
             * 给出代码提供思路
             * *************************************/
            //将文件读入dt，数据准备部分
            string filePath = null;
            Array excelspaths = excelspath.ToArray();
            DataTable dt = new DataTable();
            FileStream fsRead = null;
            IWorkbook wkBook = null;
            DataColumn column = null;
            DataRow dataRow = null;
            for (int count = 0; count < excelspaths.Length; count++)
            {
                filePath = excelspath[count];
                if (filePath != null)
                {
                    fsRead = new FileStream(filePath, FileMode.Open);//1、创建一个工作簿workBook对象
                    if (filePath.IndexOf(".xlsx") > 0)////将人员表.xls中的内容读取到fsRead中
                        wkBook = new XSSFWorkbook(fsRead);
                    else if (filePath.IndexOf(".xls") > 0)
                        wkBook = new HSSFWorkbook(fsRead);

                    for (int i = 0; i < wkBook.NumberOfSheets; i++)
                    {
                        //获取每个工作表对象
                        ISheet sheet = wkBook.GetSheetAt(i);
                        //获取每个工作表的行
                        for (int r = 0; r <= sheet.LastRowNum; r++)
                        {
                            IRow currentRow = sheet.GetRow(r);
                            if (currentRow == null) continue;
                            if (r != 0)
                            {
                                dataRow = dt.NewRow();
                            }
                            for (int c = 0; c < currentRow.LastCellNum; c++)
                            {
                                try
                                {
                                    //获取每个单元格
                                    ICell cell = currentRow.GetCell(c);
                                    if (r == 0 && count == 0)//第一次读把首行读进dt里 
                                    {
                                        //本行都存入数据结构中
                                        if (cell == null)
                                        {
                                            column = new DataColumn(" ");
                                        }
                                        else
                                        {
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    column = new DataColumn(" ");
                                                    break;
                                                case CellType.String:
                                                    column = new DataColumn(cell.StringCellValue);
                                                    break;
                                            }
                                        }
                                        dt.Columns.Add(column);
                                    }//没问题了

                                    if (r != 0)
                                    {
                                        //本行的存入
                                        if (cell == null)
                                        {
                                            dataRow[c] = " ";
                                        }
                                        else
                                        {
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[c] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[c] = cell.DateCellValue;
                                                    else
                                                        dataRow[c] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                        dataRow[c] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    ;
                                }
                            }
                            if (dataRow != null)
                            {
                                dt.Rows.Add(dataRow.ItemArray);
                                dataRow = null;
                            }
                        }
                    }
                    fsRead.Close();
                    wkBook.Close();
                    fsRead.Dispose();
                }//endif
            }

            //数据库操作部分
            /************自行修改此项***************/
            String connStr = "Data Source =.; Initial Catalog = MyQQ; User ID = sa; Pwd = sa";//本地连接 请自行修改此项！！！！！！！！！！！！！！！
            /************自行修改此项**************/

            SqlConnection sConn = new SqlConnection(connStr);

            try
            {
                sConn.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误：" + ex.Message);
            }


            SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connStr, SqlBulkCopyOptions.UseInternalTransaction);
            sqlbulkcopy.DestinationTableName = "Table_1";//数据库中的表名
            sqlbulkcopy.WriteToServer(dt);

            dt.Dispose();
            sConn.Dispose();
            // bool b =
            //String connectionString=
            // SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction);// DBHelper.Update("");
            /* if (b)
             {
                 System.Windows.Forms.MessageBox.Show("保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 //dgv.DataSource = null;
             }
             else
             {
                 System.Windows.Forms.MessageBox.Show("保存失败！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }*/

        }
    }
}
