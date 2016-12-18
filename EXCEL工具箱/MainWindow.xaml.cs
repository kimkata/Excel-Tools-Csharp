using System;
using System.Windows;
using System.IO;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Threading;
using System.Text.RegularExpressions;

namespace WpfApplication1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private string mailAddress = "";
        private string mailUsername = "";
        private string password = "";
        private string smtpServer = "";
        private string resultFile = "";
        private string xlsSheetName = "";
        private string columnNumber = "";
        private int startRowNumber = 0;
        public static bool checkBox_SplitIsChecked = false;
        public static bool checkBox_viewSendResultIsChecked = false;
        public static bool checkbox_useExistedListIsChecked = false;
        public string resultFolder = "";
        private string xls_sheet_copy = "";
        // 以下是发送邮件需要用到的成员变量
        private ArrayList fileArray = new ArrayList();
        private ArrayList fileFullnameArray = new ArrayList();
        private String mailExcel = "";
        //以下是多表拆分需要的成员变量：

        private string xls_file_for_mult_split;

        public MainWindow()
        {
            InitializeComponent();
            //在初始化窗口时，同时恢复用户名、密码、邮箱
            mailUsername = Properties.Settings.Default.mailUsername;
            textBox_username.Text = mailUsername;
            mailAddress = Properties.Settings.Default.mailAddress;
            textBox_mailAddress_input.Text = mailAddress;
            password = Properties.Settings.Default.password;
            passwordBox.Password = password;
            smtpServer = Properties.Settings.Default.smtpServer;
            textBox_smtpServer_input.Text = smtpServer;

        }

        //#############################################################################################################
        //*以下是拆分工作表的代码
        //*作者：陈志刚
        //*时间：2016-10-24
        //*############################################################################################################

        private void tabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void richTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        //以下是选择要拆分的EXCEL文件的对话框

        private void button_xlsFile_choose_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.InitialDirectory = "c:\\";
            ofd.Filter = "EXCEL文件(*.xls*)|*.xls*|所有文件(*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == true)
            {
                resultFile = ofd.FileName;
                button_xlsFile_choose.Foreground = Brushes.Blue;
                richTextBox_split.Foreground = Brushes.Blue;
                setText_RichTextBox(richTextBox_split, "您选择的文件是：" + "\n" + resultFile, Brushes.Blue, CLEAR);
                ArrayList worksheet_list = new ArrayList();
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.DisplayAlerts = false;
                _Workbook xlsWorkbook = xlApp.Workbooks.Open(resultFile);
                Sheets xlsWorksheets = xlsWorkbook.Worksheets;
                string sheets_set = "";
                foreach (_Worksheet work_sheet in xlsWorksheets)
                {
                    worksheet_list.Add(work_sheet.Name);
                    sheets_set = sheets_set + "\n" + work_sheet.Name;
                }
                setText_RichTextBox(richTextBox_split, "工作表列表如下：" + sheets_set, Brushes.Blue, CLEAR);
                comboBox_split.ItemsSource = worksheet_list;
                comboBox_copy.ItemsSource = worksheet_list;
                xlsWorkbook.Close();
                xlApp.Quit();

            }
            else
            {
                setText_RichTextBox(richTextBox_split, "您取消了选择文件！" + "\n" + resultFile, Brushes.Red, CLEAR);
            }
        }



        //以下是输入拆分文件sheet名的文本框
        /*  private void textBox_sheetName_input_TextChanged(object sender, TextChangedEventArgs e)
          {
              xlsSheetName = textBox_sheetName_input.Text;
              if (!(xlsSheetName == ""))
              {
                  setText_RichTextBox(richTextBox_split, "提醒：您输入的工作表名称是： " + xlsSheetName + "   请确保是正确的，否则会拆分失败", Brushes.Blue, CLEAR);
              }
          }
          */


        private void comboBox_split_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            xlsSheetName = (string)comboBox_split.SelectedValue;
            setText_RichTextBox(richTextBox_split, "提醒：选择的工作表名称是： \n" + xlsSheetName + "\n请确保是正确的，否则会拆分失败", Brushes.Red, CLEAR);
        }
        private void comboBox_copy_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            xls_sheet_copy = (string)comboBox_copy.SelectedValue;
            setText_RichTextBox(richTextBox_split, "提醒：选择的工作表名称是： \n" + xls_sheet_copy + "\n请确保是正确的，否则会拆分失败", Brushes.Red, CLEAR);
        }
        //以下是输入拆分文件所依据的列号文本框
        private void textBox_columnLetter_input_TextChanged(object sender, TextChangedEventArgs e)
        {
            columnNumber = textBox_columnLetter_input.Text;
            if (!(columnNumber == ""))
            {
                setText_RichTextBox(richTextBox_split, "提醒：您输入的工作表拆分列是： ", Brushes.Blue, CLEAR);
                setText_RichTextBox(richTextBox_split, columnNumber + "   列", Brushes.Red, NOTCLEAR);
                setText_RichTextBox(richTextBox_split, "请确保是正确的，否则会失败", Brushes.Blue, NOTCLEAR);
            }
        }
        //以下是输入拆分文件正文开始的行号文本框
        private void textBox_rowStartNum_input_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                startRowNumber = System.Convert.ToInt32(textBox_rowStartNum_input.Text);
                setText_RichTextBox(richTextBox_split, "正确：您输入的是数字", Brushes.Blue, CLEAR);
            }
            catch (Exception error)
            {
                setText_RichTextBox(richTextBox_split, "发生异常，信息如下：\nMessage:" + error.Message + "\nStackTrace" + error.StackTrace, Brushes.Red, CLEAR);
            }
        }


        //开始拆分按钮
        private void button_splitStart_Click(object sender, RoutedEventArgs e)
        {
            checkBox_SplitIsChecked = (bool)checkBox_Split.IsChecked;
            Thread thread = new Thread(doSplit);
            thread.Start();
        }
        public void doSplit()
        {

            ExcelSplit excelSplit = new ExcelSplit(resultFile, xlsSheetName, xls_sheet_copy, columnNumber, startRowNumber, progressbar_split, checkBox_Split, richTextBox_split);
            try
            {
                excelSplit.split();
            }
            catch (Exception error)
            {
                setText_RichTextBox(richTextBox_split, "发生异常，信息如下：\nMessage:" + error.Message + "\nStackTrace" + error.StackTrace, Brushes.Red, CLEAR);
            }

        }

        //以下方法用于设置RichTextBox显示的text内容
        const int CLEAR = 0;
        const int NOTCLEAR = 1;
        public void setText_RichTextBox(System.Windows.Controls.RichTextBox richTextBox, string text, SolidColorBrush color, int clearMode)
        {
            System.Windows.Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
            {
                if (clearMode == CLEAR)
                {
                    richTextBox.Document.Blocks.Clear();
                    richTextBox.Foreground = color;
                    richTextBox.AppendText(text);
                }
                if (clearMode == NOTCLEAR)
                {
                    richTextBox.Foreground = color;
                    richTextBox.AppendText(text);
                }

            });
        }
        /*
       #############################################################################################################
       #以下是发送邮件的代码
       #作者：陈志刚
       #时间：2016-10-24
       ############################################################################################################
       */

        //以下是发送邮件的代码click事件，重新开一个新的线程来完成发送。
        private void button_send_Click(object sender, RoutedEventArgs e)
        {
            checkBox_viewSendResultIsChecked = (bool)checkBox_viewSendResult.IsChecked;
            Thread thread_sendmail = new Thread(sendmail);
            thread_sendmail.Start();
        }
        //以下是发送邮件的方法
        public void sendmail()
        {
            try
            {
                Sendmail sendmail = new Sendmail(mailExcel, mailUsername, mailAddress, password, smtpServer, progressbar_sendmail, richTextBox_sendmail, 25);
                sendmail.checkbox_ischecked(checkBox_useExistedList, checkBox_viewSendResult);
                sendmail.send();
            }
            catch (Exception error)
            {

                setText_RichTextBox(richTextBox_sendmail, "发生异常，信息如下：\nMessage:" + error.Message + "\nStackTrace" + error.StackTrace, Brushes.Red, CLEAR);
            }

        }




        private void textBox_username_TextInput(object sender, TextCompositionEventArgs e)
        {
            mailUsername = (string)textBox_username.Text;

            if (passwordSaved.IsChecked == true)
            {
                Properties.Settings.Default.mailUsername = mailUsername;
                Properties.Settings.Default.Save();
            }
        }
        //保存用户名
        private void textBox_username_TextChanged(object sender, TextChangedEventArgs e)
        {
            mailUsername = (string)textBox_username.Text;
            if (passwordSaved.IsChecked == true)
            {
                Properties.Settings.Default.mailUsername = mailUsername;
                Properties.Settings.Default.Save();
            }
        }

        //保存邮件地址
        private void textBox_mailAddress_input_TextChanged(object sender, TextChangedEventArgs e)
        {
            mailAddress = (string)textBox_mailAddress_input.Text;

            if (passwordSaved.IsChecked == true)
            {
                Properties.Settings.Default.mailAddress = mailAddress;
                Properties.Settings.Default.Save();
            }

        }
        //保存密码
        private void passwordBox_TextInput(object sender, TextCompositionEventArgs e)
        {
            password = (string)passwordBox.Password;
            if (passwordSaved.IsChecked == true)
            {
                Properties.Settings.Default.password = password;
                Properties.Settings.Default.Save();

            }
        }

        //保存信息复选框代码
        private void passwordSaved_Click(object sender, RoutedEventArgs e)
        {
            if (passwordSaved.IsChecked == true)
            {
                Properties.Settings.Default.mailUsername = mailUsername;
                Properties.Settings.Default.mailAddress = mailAddress;
                Properties.Settings.Default.password = password;
                Properties.Settings.Default.smtpServer = smtpServer;
                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.mailUsername = "";
                Properties.Settings.Default.mailAddress = "";
                Properties.Settings.Default.password = "";
                Properties.Settings.Default.smtpServer = "";
                Properties.Settings.Default.Save();
            }

        }
        //保存SMTP
        private void textBox_smtpServer_input_TextChanged(object sender, TextChangedEventArgs e)
        {
            smtpServer = (string)textBox_smtpServer_input.Text;

            if (passwordSaved.IsChecked == true)
            {
                Properties.Settings.Default.smtpServer = smtpServer;
                Properties.Settings.Default.Save();
            }

        }
        //针对密码框修改的情况
        private void passwordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            password = (string)passwordBox.Password;
            if (passwordSaved.IsChecked == true)
            {
                Properties.Settings.Default.password = password;
                Properties.Settings.Default.Save();
            }

        }
        //选择“即将发送的文件”所在的文件夹，这里使用了windows.forms里面的控件
        private void button_folderChoose_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                resultFolder = fbd.SelectedPath;
                button_folderChoose.Foreground = System.Windows.Media.Brushes.Red;
                richTextBox_sendmail.Foreground = System.Windows.Media.Brushes.Red;
                setText_RichTextBox(richTextBox_sendmail, "您选择的文件夹是：" + "\n" + resultFolder, Brushes.Blue, CLEAR);
            }
            else
            {
                setText_RichTextBox(richTextBox_sendmail, "您取消了选择文件夹", Brushes.Red, CLEAR);
            }

        }

        //#############################################################################################################################################################################
        //#############################################################################################################################################################################
        //#############################################################################################################################################################################
        //检索文件夹下的文件方法
        private void button_fileFind_Click(object sender, RoutedEventArgs e)
        {
            checkbox_useExistedListIsChecked = (bool)checkBox_useExistedList.IsChecked;
            checkFileList();
        }

        void WalkDirectoryTree(System.IO.DirectoryInfo root)
        {

            System.IO.FileInfo[] files = null;
            System.IO.DirectoryInfo[] subDirs = null;
            try
            {
                files = root.GetFiles("*.*");
            }
            catch (UnauthorizedAccessException)
            {

            }

            catch (System.IO.DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }

            if (files != null)
            {
                foreach (System.IO.FileInfo fi in files)
                {
                    fileArray.Add(fi.Name);
                    fileFullnameArray.Add(fi.FullName);
                }

                subDirs = root.GetDirectories();

                foreach (System.IO.DirectoryInfo dirInfo in subDirs)
                {
                    WalkDirectoryTree(dirInfo);
                }
            }
        }

        public void checkFileList()
        {
            try
            {
                if (checkbox_useExistedListIsChecked == false)
                {
                    int k;
                    DirectoryInfo theFolder = new DirectoryInfo(resultFolder);
                    WalkDirectoryTree(theFolder);

                    //############################将检索到的信息填写到EXCEL表中##############################################################################

                    k = fileArray.Count;
                    setText_RichTextBox(richTextBox_sendmail, "1.文件检索完毕:\n", Brushes.Blue, CLEAR);
                    Microsoft.Office.Interop.Excel.Application xlApp2 = new Microsoft.Office.Interop.Excel.Application();
                    xlApp2.DisplayAlerts = false;
                    _Workbook xlsWorkbook2 = xlApp2.Workbooks.Add();
                    _Worksheet xlsWorksheet2 = xlsWorkbook2.Sheets["sheet1"];
                    setText_RichTextBox(richTextBox_sendmail, "\n2.正在调用Microsoft Excel程序..." + "\n", Brushes.Blue, NOTCLEAR);
                    xlsWorksheet2.Cells[1, 1] = "要发送的附件列表如下，请补全剩余信息：";
                    xlsWorksheet2.Cells[2, 1] = "收件人姓名";
                    xlsWorksheet2.Cells[2, 2] = "收件人邮件地址";
                    xlsWorksheet2.Cells[2, 3] = "邮件主题";
                    xlsWorksheet2.Cells[2, 4] = "邮件正文";
                    xlsWorksheet2.Cells[2, 5] = "vlookup函数辅助列";
                    xlsWorksheet2.Cells[2, 5].Interior.ColorIndex = 4;
                    xlsWorksheet2.Cells[2, 6] = "要发送的附件路径（本列不可修改）";
                    setText_RichTextBox(richTextBox_sendmail, "3.正在生成需要填写的邮件附件信息" + "\n", Brushes.Blue, NOTCLEAR);
                    for (int i = 0; i < k; i++)
                    {
                        // xlsWorksheet2.Cells[i + 3, 3] = fileArray[i];
                        //  xlsWorksheet2.Cells[i + 3, 4] = fileArray[i];
                        xlsWorksheet2.Cells[i + 3, 5] = fileArray[i];
                        xlsWorksheet2.Cells[i + 3, 6] = fileFullnameArray[i];
                    }

                    if (checkBox_replace.IsChecked == true)
                    {
                        xlsWorksheet2.Columns["E:E"].Replace("*_", "", XlLookAt.xlPart, XlSearchOrder.xlByRows, false, false, false);
                        xlsWorksheet2.Columns["E:E"].Replace(".*", "", XlLookAt.xlPart, XlSearchOrder.xlByRows, false, false, false);
                    }

                    mailExcel = theFolder + "\\" + k + ".生成的邮件附件列表.xlsx";
                    xlsWorkbook2.SaveAs(mailExcel);
                    xlsWorkbook2.Close();
                    xlApp2.Quit();
                    setText_RichTextBox(richTextBox_sendmail, "4.检索附件（检索完毕）,请在即将打开的EXCEL文件中填写详细的邮件发送信息", Brushes.Blue, NOTCLEAR);
                    System.Diagnostics.Process.Start(mailExcel);
                }

                else
                {
                    System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
                    ofd.InitialDirectory = "c:\\";
                    ofd.Filter = "EXCEL文件(*.xls*)|*.xls*|所有文件(*.*)|*.*";
                    ofd.FilterIndex = 1;
                    ofd.RestoreDirectory = true;
                    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        mailExcel = ofd.FileName;
                        setText_RichTextBox(richTextBox_sendmail, "您选择的邮件列表加载文件是：" + "\n", Brushes.Blue, CLEAR);
                        setText_RichTextBox(richTextBox_sendmail, mailExcel, Brushes.Blue, NOTCLEAR);
                    }
                }
            }
            catch (Exception error)
            {
                setText_RichTextBox(richTextBox_sendmail, "发生异常，信息如下：\nMessage:" + error.Message + "\nStackTrace" + error.StackTrace, Brushes.Red, CLEAR);
            }
        }
        //#############################################################################################################################################################################
        //#############################################################################################################################################################################
        //#############################################################################################################################################################################


        //使用已经存在的邮件发送列表对应的代码
        private void checkBox_useExistedList_Click(object sender, RoutedEventArgs e)
        {
            if (checkBox_useExistedList.IsChecked == true)
            {
                button_fileFind.Foreground = Brushes.Blue;
                button_fileFind.Content = "->点击此处\n选择邮件信息所在的Excel表格！";
                setText_RichTextBox(richTextBox_sendmail, "已经切换为从【已存在的Excel表格】直接发送邮件!", Brushes.Blue, CLEAR);
            }
            else
            {
                button_fileFind.Foreground = Brushes.Black;
                button_fileFind.Content = "检索文件";
                setText_RichTextBox(richTextBox_sendmail, "请点击按钮，选择要发送的附件所在的文件夹", Brushes.Blue, CLEAR);
            }
        }



        private void textBox_mailAddress_input_TextInput(object sender, TextCompositionEventArgs e)
        {


        }

        private void textBox_mailAddress_input_LostFocus(object sender, RoutedEventArgs e)
        {
            mailAddress = (string)textBox_mailAddress_input.Text;
            Regex regMailAddress = new Regex("[\\w!#$%&'*+/=?^_`{|}~-]+(?:\\.[\\w!#$%&'*+/=?^_`{|}~-]+)*@(?:[\\w](?:[\\w-]*[\\w])?\\.)+[\\w](?:[\\w-]*[\\w])?");
            Regex regMailAddressType = new Regex("[a-zA-Z0-9]+(?=.(com|net))");
            bool isMailAddress = regMailAddress.IsMatch(mailAddress);
            if (!isMailAddress)
            {

                setText_RichTextBox(richTextBox_sendmail, "出现错误,可能原因：\n" +
                                                           "邮箱地址填写不规范，请仔细检查，正确的邮箱地址格式为：用户名@邮箱域名，例如  chen@qq.com\n", Brushes.Red, CLEAR);


            }
            else
            {
                string mailAddressType = regMailAddressType.Match(mailAddress).ToString();
                setText_RichTextBox(richTextBox_sendmail, "恭喜您，您使用的是" + mailAddressType + "邮箱，邮箱填写是规范的！\n" +
                                                          "请确保发送邮件的邮箱开启了smtp功能，详细操作流程可以咨询邮箱管理员。\n" +
                                                          "另外，大量发送邮件，可能会被误判为发送垃圾邮件，被邮件服务器拦截。", Brushes.Blue, CLEAR);
            }
        }


        private string xlsWorkbook_setup_path;
        private void button_make_split_info_Click(object sender, RoutedEventArgs e)
        {
            if (checkBox_mult_split_ischecked == false)
            {
                try
                {
                    ArrayList worksheet_list = new ArrayList();
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlApp.DisplayAlerts = false;
                    _Workbook xlsWorkbook = xlApp.Workbooks.Open(xls_file_for_mult_split);
                    _Workbook xlsWorkbook_setup = xlApp.Workbooks.Add();
                    _Worksheet xlsWorksheet_setup_first = xlsWorkbook_setup.Worksheets.Add();
                    xlsWorksheet_setup_first.Name = "参数设置";
                    _Worksheet xlsWorksheet_setup_second = xlsWorkbook_setup.Worksheets.Add();
                    xlsWorksheet_setup_second.Name = "成员别名";
                    Sheets xlsWorksheets = xlsWorkbook.Worksheets;

                    foreach (_Worksheet work_sheet in xlsWorksheets)
                    {
                        worksheet_list.Add(work_sheet.Name);
                    }


                    for (int i = 0; i < worksheet_list.Count; i++)
                    {

                        xlsWorksheet_setup_first.Cells[i + 2, 1] = worksheet_list[i];
                    }
                    xlsWorksheet_setup_first.Cells[1, 1] = "工作表名称";
                    xlsWorksheet_setup_first.Cells[1, 2] = "是参与拆分（1是，0否，-1否，且复制到新表）";
                    xlsWorksheet_setup_first.Cells[1, 3] = "拆分列号";
                    xlsWorksheet_setup_first.Cells[1, 4] = "正文开始行号";
                    xlsWorksheet_setup_first.Columns["A:D"].EntireColumn.AutoFit();
                    xlsWorksheet_setup_second.Cells[1, 1] = "成员名";
                    xlsWorksheet_setup_second.Cells[1, 2] = "成员别名";
                    xlsWorksheet_setup_second.Columns["A:B"].EntireColumn.AutoFit();
                    xlsWorkbook_setup.Sheets["Sheet1"].delete();
                    xlsWorksheet_setup_first.Activate();
                    xlsWorkbook_setup_path = xlsWorkbook.Path + "\\工作表拆分参数列表.xlsx";
                    xlsWorkbook_setup.SaveAs(xlsWorkbook_setup_path);
                    xlsWorkbook.Close();
                    xlsWorkbook_setup.Close();
                    xlApp.Quit();
                    setText_RichTextBox(richTextBox_mult_split, "\n已经生成参数列表文件：" + "\n" + xlsWorkbook_setup_path, Brushes.Blue, NOTCLEAR);
                    setText_RichTextBox(richTextBox_mult_split, "\n正在打开工作表<<工作表拆分参数列表.xlsx>>,请填写完毕后，保存并关闭", Brushes.Blue, NOTCLEAR);
                    System.Diagnostics.Process.Start(xlsWorkbook_setup_path);

                }
                catch (Exception error)
                {
                    setText_RichTextBox(richTextBox_mult_split, "发生异常，信息如下：\nMessage:" + error.Message + "\nStackTrace:" + error.StackTrace, Brushes.Red, CLEAR);

                }
            }
            else
            {
                System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
                ofd.InitialDirectory = "c:\\";
                ofd.Filter = "EXCEL文件(*.xls*)|*.xls*|所有文件(*.*)|*.*";
                ofd.FilterIndex = 1;
                ofd.RestoreDirectory = true;
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    xlsWorkbook_setup_path = ofd.FileName;
                    setText_RichTextBox(richTextBox_mult_split, "您选择的加载文件是：" + "\n", Brushes.Blue, CLEAR);
                    setText_RichTextBox(richTextBox_mult_split, xlsWorkbook_setup_path, Brushes.Blue, NOTCLEAR);
                }
            }
        }


        bool checkBox_mult_split_ischecked = false;
        private void checkBox_mult_split_Click(object sender, RoutedEventArgs e)
        {
            if (checkBox_mult_split.IsChecked == true)
            {
                button_make_split_info.Foreground = Brushes.Blue;
                button_mult_xlsfile_choose.Foreground = Brushes.Blue;
                button_make_split_info.Content = "->点击此处\n选择列表所在表格！";
                button_mult_xlsfile_choose.Content = "->点击此处\n选择要拆分的表格！";

                checkBox_mult_split_ischecked = true;
                setText_RichTextBox(richTextBox_mult_split, "注意：要拆分的文件和参数列表文件必须同时选择！\n注意：要拆分的文件和参数列表文件必须同时选择！\n已经切换为从【已存在的Excel表格】执行!", Brushes.Blue, CLEAR);
            }
            else
            {
                checkBox_mult_split_ischecked = false;
                button_make_split_info.Foreground = Brushes.Black;
                button_make_split_info.Content = "生成拆分参数列表";
                button_mult_xlsfile_choose.Foreground = Brushes.Black;
                button_mult_xlsfile_choose.Content = "文件浏览";
                setText_RichTextBox(richTextBox_mult_split, "请选择要拆分的工作表", Brushes.Blue, CLEAR);

            }
        }

        //################################################################################################################################################
        //################################################################################################################################################
        //################################################################################################################################################
        //多表拆分的方法
        private void button_mult_split_Click(object sender, RoutedEventArgs e)
        {
            Thread thread = new Thread(excelmultsplit);
            thread.Start();


        }


        void excelmultsplit()
        {

            ArrayList array_sheet_split = new ArrayList();// array_sheet_split是列表的列表，用于存放xlsworksheet中的拆分信息。
            const int SPLIT_OK = 1;
            const int SPLIT_NO_BUT_COPY = -1;
            Microsoft.Office.Interop.Excel.Application xlsxApp = new Microsoft.Office.Interop.Excel.Application();
            xlsxApp.DisplayAlerts = false;
            _Workbook xlsWorkbook_Info = xlsxApp.Workbooks.Open(xlsWorkbook_setup_path);
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
                setText_RichTextBox(richTextBox_mult_split, "已经生成文件：" + workbookName, Brushes.Blue, CLEAR);
                newBook.Sheets["Sheet1"].Delete();
                newBook.SaveAs(workbookName);
                newBook.Close();
                System.Windows.Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                {
                    progressbar_mult_split.Maximum = progressbar_maxvalue;
                    progressbar_mult_split.Value = progressbar_value++;
                });



            }





            xlsWorkbook_Info.Close();
            xlsWorkbook_split.Close();
            xlsxApp.Quit();
            setText_RichTextBox(richTextBox_mult_split, "程序运行完毕,请查看相关目录的文件夹：\n" + myPath, Brushes.Blue, CLEAR);

            //以下方法打开工作表目录会被360安全卫士拦截
            //  System.Diagnostics.Process.Start("explorer.exe", myPath);

        }




        //#####################################################################################################################################################
        //#####################################################################################################################################################
        //#####################################################################################################################################################
        //以下是批量重命名文件的模块

        private string folder_check = "";
        private void button_choose_folder_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folder_check = fbd.SelectedPath;
                button_folderChoose.Foreground = System.Windows.Media.Brushes.Red;
                richTextBox_sendmail.Foreground = System.Windows.Media.Brushes.Red;
                setText_RichTextBox(richTextBox_rename, "您选择的文件夹是：" + "\n" + folder_check, Brushes.Blue, CLEAR);
            }
            else
            {
                setText_RichTextBox(richTextBox_rename, "您取消了选择文件夹！", Brushes.Red, CLEAR);
            }
        }


        private ArrayList chekfileArray = new ArrayList();
        private ArrayList checkfileFullnameArray = new ArrayList();
        private string excel_check;
        private void button_check_file_Click(object sender, RoutedEventArgs e)
        {
            checkFileList_rename();

        }



        void WalkDirectoryTree_check(System.IO.DirectoryInfo root)
        {

            System.IO.FileInfo[] files = null;
            System.IO.DirectoryInfo[] subDirs = null;
            try
            {
                files = root.GetFiles("*.*");
            }
            //  catch (UnauthorizedAccessException)
            //   {

            //   }

            catch (System.IO.DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }

            if (files != null)
            {
                foreach (System.IO.FileInfo fi in files)
                {
                    chekfileArray.Add(fi.Name);
                    checkfileFullnameArray.Add(fi.FullName);
                }

                subDirs = root.GetDirectories();

                foreach (System.IO.DirectoryInfo dirInfo in subDirs)
                {
                    WalkDirectoryTree_check(dirInfo);
                }
            }
        }

        public void checkFileList_rename()
        {
            try
            {
                if (checkbox_rename_ischecked == false)
                {
                    int k;
                    DirectoryInfo theFolder = new DirectoryInfo(folder_check);
                    WalkDirectoryTree_check(theFolder);

                    //############################将检索到的信息填写到EXCEL表中##############################################################################

                    k = chekfileArray.Count;
                    setText_RichTextBox(richTextBox_rename, "1.文件检索完毕:\n", Brushes.Blue, CLEAR);
                    Microsoft.Office.Interop.Excel.Application xlApp2 = new Microsoft.Office.Interop.Excel.Application();
                    xlApp2.DisplayAlerts = false;
                    _Workbook xlsWorkbook2 = xlApp2.Workbooks.Add();
                    _Worksheet xlsWorksheet2 = xlsWorkbook2.Sheets["sheet1"];
                    setText_RichTextBox(richTextBox_sendmail, "\n2.正在调用Microsoft Excel程序..." + "\n", Brushes.Blue, NOTCLEAR);
                    xlsWorksheet2.Cells[1, 1] = "您选择的文件夹内的文件列表如下：请补全剩余信息：";
                    xlsWorksheet2.Cells[2, 1] = "修改后的文件名（注：包括后缀名，类似于C列）";
                    xlsWorksheet2.Cells[2, 2] = "修改后的文件全名（注：包含路径，类似于D列）";
                    xlsWorksheet2.Cells[2, 3] = "文件名";
                    xlsWorksheet2.Cells[2, 4] = "文件全名";
                    for (int i = 0; i < k; i++)
                    {
                        xlsWorksheet2.Cells[i + 3, 3] = chekfileArray[i];
                        xlsWorksheet2.Cells[i + 3, 4] = checkfileFullnameArray[i];
                    }
                    xlsWorksheet2.Columns["A:B"].autofit();

                    excel_check = theFolder + "\\" + k + ".生成的文件批量重新命名列表.xlsx";
                    xlsWorkbook2.SaveAs(excel_check);
                    xlsWorkbook2.Close();
                    xlApp2.Quit();
                    setText_RichTextBox(richTextBox_rename, "\n3.检索附件（检索完毕）,请在即将打开的EXCEL文件中填写详细信息", Brushes.Blue, NOTCLEAR);
                    System.Diagnostics.Process.Start(excel_check);
                }

                else
                {
                    Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
                    ofd.InitialDirectory = "c:\\";
                    ofd.Filter = "EXCEL文件(*.xls*)|*.xls*|所有文件(*.*)|*.*";
                    ofd.FilterIndex = 1;
                    ofd.RestoreDirectory = true;
                    if (ofd.ShowDialog() == true)
                    {
                        excel_check = ofd.FileName;
                        setText_RichTextBox(richTextBox_rename, "您选择的文件是：" + "\n" + excel_check, Brushes.Blue, CLEAR);

                    }
                }
            }
            catch (Exception error)
            {
                setText_RichTextBox(richTextBox_rename, "发生异常，信息如下：\nMessage:" + error.Message + "\nStackTrace" + error.StackTrace, Brushes.Red, CLEAR);
            }
        }



        bool checkbox_rename_ischecked = false;
        private void checkBox_rename_Checked(object sender, RoutedEventArgs e)
        {
            //  checkbox_rename_ischecked = true;
        }

        private void checkBox_rename_Click(object sender, RoutedEventArgs e)
        {
            if (checkBox_rename.IsChecked == true)
            {
                button_check_file.Foreground = Brushes.Blue;
                button_check_file.Content = "->点击此处\n选择相关的Excel表格！";
                setText_RichTextBox(richTextBox_rename, "已经切换为从【已存在的Excel表格】执行!", Brushes.Blue, CLEAR);
                checkbox_rename_ischecked = true;

            }
            else
            {
                button_check_file.Foreground = Brushes.Black;
                button_check_file.Content = "检索文件";
                setText_RichTextBox(richTextBox_rename, "请点击按钮，选择要发送的附件所在的文件夹", Brushes.Blue, CLEAR);
                checkbox_rename_ischecked = false;
            }
        }

        private void button_start_rename_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                ArrayList arrayList_old_file_name = new ArrayList();
                ArrayList arrayList_new_file_name = new ArrayList();
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.DisplayAlerts = false;
                _Workbook xlsWorkbook = xlApp.Workbooks.Open(excel_check);
                _Worksheet xlsWorksheet = xlsWorkbook.Sheets[1];
                int row_nums = xlsWorksheet.UsedRange.Rows.Count;
                int columns = xlsWorksheet.UsedRange.Columns.Count;
                setText_RichTextBox(richTextBox_rename, excel_check, Brushes.Blue, CLEAR);


                for (int i = 0; i < row_nums; i++)
                {
                    string old_file_name = "";
                    string new_file_name = "";
                    new_file_name = System.Convert.ToString(xlsWorksheet.Cells[i + 3, 2].value);
                    old_file_name = System.Convert.ToString(xlsWorksheet.Cells[i + 3, 4].value);
                    if (new_file_name != "" && old_file_name != "" && new_file_name != old_file_name)
                    {
                        arrayList_new_file_name.Add(new_file_name);
                        arrayList_old_file_name.Add(old_file_name);
                    }

                }

                for (int i = 0; i < arrayList_old_file_name.Count; i++)
                {

                    if (File.Exists(System.Convert.ToString(arrayList_old_file_name[i])))
                    {

                        FileInfo fileInfo = new FileInfo(System.Convert.ToString(arrayList_old_file_name[i]));
                        fileInfo.MoveTo(System.Convert.ToString(arrayList_new_file_name[i]));
                    }

                }
                xlsWorkbook.Close();
                xlApp.Quit();
                setText_RichTextBox(richTextBox_rename, "程序运行完毕,请查看相关的文件夹", Brushes.Blue, CLEAR);
            }

            catch (Exception error)
            {

                setText_RichTextBox(richTextBox_rename, error.Message + "\n" + error.StackTrace, Brushes.Blue, CLEAR);
            }

        }



        private void button_mult_xlsfile_choose_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.InitialDirectory = "c:\\";
            ofd.Filter = "EXCEL文件(*.xls*)|*.xls*|所有文件(*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == true)
            {
                xls_file_for_mult_split = ofd.FileName;
                button_mult_xlsfile_choose.Foreground = Brushes.Blue;
                setText_RichTextBox(richTextBox_mult_split, "您选择的文件是：" + "\n" + xls_file_for_mult_split, Brushes.Blue, CLEAR);
            }
            else
            {
                setText_RichTextBox(richTextBox_mult_split, "您取消了选择文件！" + "\n" + resultFile, Brushes.Red, CLEAR);
            }
        }


    }
}


