using System;
using System.IO;
using System.Windows.Threading;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Windows.Media;

namespace WpfApplication1
{
    //以上是由陈志刚编写的类
    //以下是EXCEL拆分类，作者：陈志刚

    public class ExcelSplit
    {
        private string xlsFullName;
        private string xlsSheetName;
        private string xlssheet_copy;
        private string columnNumber;
        private int startRowNumber = 0;
        System.Windows.Controls.CheckBox checkBox_Split;
        System.Windows.Controls.RichTextBox richTextBox;
        System.Windows.Controls.ProgressBar progressbar;
        //以下是ExcelSplit的构造函数
        public ExcelSplit(string xlsFullName, string xlsSheetName, string xlssheet_copy, string columnNumber, int startRowNumber, System.Windows.Controls.ProgressBar progressbar, System.Windows.Controls.CheckBox checkBox_Split, System.Windows.Controls.RichTextBox richTextBox)
        {
            this.xlsFullName = xlsFullName;
            this.xlsSheetName = xlsSheetName;
            this.columnNumber = columnNumber;
            this.startRowNumber = startRowNumber;
            this.checkBox_Split = checkBox_Split;
            this.progressbar = progressbar;
            this.richTextBox = richTextBox;
            this.xlssheet_copy = xlssheet_copy;
        }

        public void progressbar_display( int value,int max_value)
        {
            System.Windows.Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
            {
                progressbar.Maximum =max_value;
                progressbar.Value = value;
            });
        }



        public void setText_RichTextBox(System.Windows.Controls.RichTextBox richTextBox, string text, SolidColorBrush color)
        {
            System.Windows.Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
            {
                richTextBox.Document.Blocks.Clear();
                richTextBox.Foreground = color;
                richTextBox.AppendText(text);
            });
        }
        //以下是split方法
        public void split()
        {
            string myPath = "";
            setText_RichTextBox(richTextBox, "正在为您拆分工作表，请耐心等待...", Brushes.Blue);
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts = false;
            _Workbook xlsWorkbook = xlApp.Workbooks.Open(xlsFullName);
            _Worksheet xlsWorksheet = (_Worksheet)xlsWorkbook.Sheets[xlsSheetName];
            _Worksheet mySheet = (_Worksheet)xlsWorkbook.Worksheets.Add();
            int k = xlsWorksheet.UsedRange.Columns.Count;
            int r = xlsWorksheet.UsedRange.Rows.Count;
            String s = columnNumber + (startRowNumber) + ":" + columnNumber + r;
            string s2 = columnNumber + startRowNumber.ToString();
            int fieldNum = mySheet.Range[columnNumber + 1].Column;

            if (xlsWorksheet.FilterMode == true)
            {
                xlsWorksheet.Rows[startRowNumber - 1].AutoFilter();
            }
            xlsWorksheet.Range[s].Copy();
            xlsWorksheet.Range[s2].PasteSpecial(XlPasteType.xlPasteValues);
            xlsWorksheet.Range[s].Copy(mySheet.Cells[1, 1]);
            mySheet.UsedRange.RemoveDuplicates(1, XlYesNoGuess.xlNo);
            int row = mySheet.Cells[1, 1].CurrentRegion.Rows.Count;
            setText_RichTextBox(richTextBox, "正在生成所需要的数组...", Brushes.Blue);
            ArrayList list = new ArrayList();
            for (int i = 1; i <= row; i++)
            {
                list.Add(System.Convert.ToString(mySheet.Cells[i, 1].Value));
            }

            mySheet.Delete();
            if (list.Contains(""))
            {
                list.Remove("");
            }

            setText_RichTextBox(richTextBox, "数组生成完毕...", Brushes.Blue);
            myPath = xlsWorkbook.Path + "\\" + xlsWorksheet.Name + "_拆分的工作表";
            if (!Directory.Exists(myPath))
            {
                DirectoryInfo dir = new DirectoryInfo(myPath);
                dir.Create();
            }


            
            int x = 0;
            int y = list.Count;
            foreach (String t in list)
            {
                String workbookName = myPath + "\\" + xlsWorksheet.Name + "_" + t;
                setText_RichTextBox(richTextBox, "正在生成工作表   " + myPath + "_" + t + ".xlsx", Brushes.Blue);
                _Workbook newBook = xlApp.Workbooks.Add();
                _Worksheet newSheet = newBook.Sheets[1];
                newSheet.Name = xlsWorksheet.Name;
                xlsWorksheet.Rows[startRowNumber - 1].AutoFilter(fieldNum, t);
                xlsWorksheet.UsedRange.Copy(newSheet.Cells[1, 1]);
                xlsWorksheet.Rows[startRowNumber - 1].AutoFilter();
                if (xlssheet_copy != "")
                {
                    xlsWorkbook.Sheets[xlssheet_copy].Copy(newSheet);
                    newSheet.Activate();
                }
                newBook.SaveAs(workbookName);
                newBook.Close();

                //显示拆分的进度
                x++;
                progressbar_display(x,y);              
               
            }
            xlsWorkbook.Close();
            xlApp.DisplayAlerts = true;
            xlApp.Quit();
            setText_RichTextBox(richTextBox, "拆分完毕，已关闭工作表，并保存到文件夹：" + myPath, Brushes.Red);
            if (MainWindow.checkBox_SplitIsChecked)
            {
                System.Diagnostics.Process.Start("explorer.exe", myPath);
            }
        }

       
        void xls_mult_split(string xls_file_for_mult_split, string xls_file_info, System.Windows.Controls.ProgressBar progressbar, System.Windows.Controls.CheckBox checkBox_Split, System.Windows.Controls.RichTextBox richTextBox)//此方法输入2个参数，一个是要拆分的工作表，另一个是存放拆分信息的工作表。
        {
            ArrayList array_sheet_split = new ArrayList();// array_sheet_split是列表的列表，用于存放xlsworksheet中的拆分信息。
            const int SPLIT_OK = 1;
            const int SPLIT_NO_BUT_COPY = -1;
            Microsoft.Office.Interop.Excel.Application xlsxApp = new Microsoft.Office.Interop.Excel.Application();
            xlsxApp.DisplayAlerts = false;
            _Workbook xlsWorkbook_Info = xlsxApp.Workbooks.Open(xls_file_info);
            _Workbook xlsWorkbook_split = xlsxApp.Workbooks.Open(xls_file_for_mult_split);
            _Worksheet xlsWorksheet_Info = xlsWorkbook_Info.Sheets["参数设置"];
            // _Worksheet xlsWorksheet2 = xlsWorkbook_Info.Sheets[1];//暂时未使用
            int num = xlsWorksheet_Info.UsedRange.Rows.Count - 1;
            ArrayList sheets_split_no_but_copy = new ArrayList();
            for (int i = 0; i < num; i++)
            {
                ArrayList arraylist = new ArrayList();
                if (System.Convert.ToInt32(xlsWorksheet_Info.Cells[i + 2, 2].Value) == SPLIT_OK)
                {
                    arraylist.Add(System.Convert.ToString(xlsWorksheet_Info.Cells[i + 2, 1].Value));
                    arraylist.Add(System.Convert.ToInt32(xlsWorksheet_Info.Cells[i + 2, 2].Value));
                    arraylist.Add(System.Convert.ToString(xlsWorksheet_Info.Cells[i + 2, 3].Value));
                    arraylist.Add(System.Convert.ToInt32(xlsWorksheet_Info.Cells[i + 2, 4].Value));
                    array_sheet_split.Add(arraylist);

                }
                else if (System.Convert.ToInt32(xlsWorksheet_Info.Cells[i + 2, 2].value) == SPLIT_NO_BUT_COPY)
                {
                    sheets_split_no_but_copy.Add(System.Convert.ToString(xlsWorksheet_Info.Cells[i + 2, 1].Value));
                }
            }
            //array_list_data是用于存放某一个表中唯一的字段成员
            ArrayList array_list_data = new ArrayList();
            foreach (ArrayList array_list in array_sheet_split)
            {
                string sheet_name = System.Convert.ToString(array_list[0]);
                string key_column = System.Convert.ToString(array_list[2]);
                int start_row = System.Convert.ToInt32(array_list[3]);
                _Worksheet xlsWorksheet_split = xlsWorkbook_split.Sheets[sheet_name];
                _Worksheet mySheet = xlsWorkbook_split.Worksheets.Add();
                int k = xlsWorksheet_split.UsedRange.Columns.Count;
                int r = xlsWorksheet_split.UsedRange.Rows.Count;
                String s = key_column + (start_row) + ":" + key_column + r;
                string s2 = key_column + System.Convert.ToString(start_row);
                int fieldNum = xlsWorksheet_split.Range[key_column + 1].Column;
                if (xlsWorksheet_split.FilterMode == true)
                {
                    xlsWorksheet_split.Rows[start_row - 1].AutoFilter();
                }
                xlsWorksheet_split.Range[s].Copy();
                xlsWorksheet_split.Range[s2].PasteSpecial(XlPasteType.xlPasteValues);
                xlsWorksheet_split.Range[s].Copy(mySheet.Cells[1, 1]);
                mySheet.UsedRange.RemoveDuplicates(1, XlYesNoGuess.xlNo);
                int row = mySheet.Cells[1, 1].CurrentRegion.Rows.Count;
                ArrayList list = new ArrayList();
                for (int i = 1; i <= row; i++)
                {
                    list.Add(System.Convert.ToString(mySheet.Cells[i, 1].Value));
                }
                mySheet.Delete();
                if (list.Contains(""))
                {
                    list.Remove("");
                }
                array_list_data.Add(list);
            }

            //创建一个包含全部字段成员的列表
            ArrayList arraylist_all_data = new ArrayList();
            foreach (ArrayList arraylist in array_list_data)
            {
                foreach (string key in arraylist)
                {
                    if (!arraylist_all_data.Contains(key))
                    {
                        arraylist_all_data.Add(key);
                    }
                }
            }
            //创建一个拆分出来的表格的保存路径
            string myPath = xlsWorkbook_split.Path + "\\" + xlsWorkbook_split.Name + "_拆分的工作表";
            if (!Directory.Exists(myPath))
            {
                DirectoryInfo dir = new DirectoryInfo(myPath);
                dir.Create();
            }

            int progressbar_value = 1;
            int progressbar_maxvalue = arraylist_all_data.Count;
            foreach (string key in arraylist_all_data)
            {
                _Workbook newBook = xlsxApp.Workbooks.Add();//每一个KEY产生一个workbook
                String workbookName = myPath + "\\" + xlsWorkbook_split.Name + "_" + key + ".xlsx";//每一个workbook的命名方式
                int i = 0;
                foreach (ArrayList array_list in array_sheet_split)//存放了需要拆分的工作表的名称信息
                {
                    string sheet_name = System.Convert.ToString(array_list[0]);//新生成的表格sheet页仍然以原有的名字来命名
                    _Worksheet sheet_for_split = xlsWorkbook_split.Sheets[sheet_name];
                    ArrayList mylist = (ArrayList)array_list_data[i];
                    i++;
                    if (mylist.Contains(key))
                    {
                        _Worksheet myNewSheet = newBook.Worksheets.Add();
                        myNewSheet.Name = sheet_name;
                        string key_column = System.Convert.ToString(array_list[2]);
                        int start_row = System.Convert.ToInt32(array_list[3]);
                        int fieldNum = myNewSheet.Range[key_column + 1].Column;
                        sheet_for_split.Rows[start_row - 1].AutoFilter(fieldNum, key);
                        sheet_for_split.UsedRange.Copy(myNewSheet.Cells[1, 1]);
                        sheet_for_split.Rows[start_row - 1].AutoFilter();

                        if (sheets_split_no_but_copy.Count != 0)
                        {
                            foreach (Object xlssheet_copy in sheets_split_no_but_copy)
                            {
                                xlsWorkbook_split.Sheets[System.Convert.ToString(xlssheet_copy)].Copy(myNewSheet);

                            }
                        }
                    }


                }
                setText_RichTextBox(richTextBox, "已经生成文件：" + workbookName, Brushes.Blue);
                newBook.Sheets["Sheet1"].Delete();
                newBook.SaveAs(workbookName);
                newBook.Close();
                System.Windows.Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                {
                    progressbar.Maximum = progressbar_maxvalue;
                    progressbar.Value = progressbar_value++;
                });



            }





            xlsWorkbook_Info.Close();
            xlsWorkbook_split.Close();
            xlsxApp.Quit();
            setText_RichTextBox(richTextBox, "程序运行完毕,请查看相关目录的文件夹：\n" + myPath, Brushes.Blue);

            //以下方法打开工作表目录会被360安全卫士拦截
            //  System.Diagnostics.Process.Start("explorer.exe", myPath);
        }

    }
}
