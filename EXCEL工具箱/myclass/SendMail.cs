using System;
using System.Collections;
using System.IO;
using System.Threading;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Diagnostics;

namespace WpfApplication1
{
   public class Sendmail
    {    //以下是发送邮件需要用到的成员变量
        private static bool checkBox_viewSendResultIsChecked = false;
        private static bool checkbox_useExistedListIsChecked = false;
        private String mailUsername = "";
        private String mailAddress = "";
        private String myPassword = "";
        private String smtpServer = "";
        private int smtpPort = 25;
        private ArrayList fileArray = new ArrayList();
        private ArrayList fileFullnameArray = new ArrayList();
        private ArrayList recepientNameArray = new ArrayList();
        private ArrayList recepientmailAddressArray = new ArrayList();
        private ArrayList msgSubjectArray = new ArrayList();
        private ArrayList msgHtmlBodyArray = new ArrayList();
        private ArrayList attachmentsName = new ArrayList();
        private String mailExcel = "";
        private RichTextBox richtextbox;
        private ProgressBar progressbar;
        private int sucessful = 0;
        private int failure = 0;
        public Sendmail(string mailExcel, string mailUsername, string mailAddress,string myPassword,string smtpServer, ProgressBar progressbar, RichTextBox richtextbox, int smtpPort = 25 )
        {
            this.mailExcel = mailExcel;
            this.mailUsername = mailUsername;
            this.mailAddress = mailAddress;
            this.myPassword = myPassword;
            this.smtpPort = smtpPort;
            this.richtextbox = richtextbox;
            this.progressbar = progressbar;

        }

        public void checkbox_ischecked(System.Windows.Controls.CheckBox checkBox_viewSendResult, System.Windows.Controls.CheckBox checkbox_useExistedList)
        {

            System.Windows.Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
            {
                checkBox_viewSendResultIsChecked = (bool)checkBox_viewSendResult.IsChecked;
                checkbox_useExistedListIsChecked = (bool)checkbox_useExistedList.IsChecked;
            });
        }

        public void progressbar_display(int value, int max_value)
        {
            System.Windows.Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
            {
                progressbar.Maximum = max_value;
                progressbar.Value = value;
            });
        }



        //以下是发送邮件的方法
        public void send()
        {
            try
            {
                //将发送邮件使用到的各类数组全部清空，防止由于在检索之后删除文件，更改列表等原因造成数组错行。
                recepientNameArray.Clear();
                recepientmailAddressArray.Clear();
                msgSubjectArray.Clear();
                msgHtmlBodyArray.Clear();
                fileFullnameArray.Clear();
                attachmentsName.Clear();
                //打开邮件发送列表mailExcel文件，读取要发送的邮件列表，将信息放置到以上数组之中。
                Microsoft.Office.Interop.Excel.Application xlsxApp = new Microsoft.Office.Interop.Excel.Application();
                xlsxApp.DisplayAlerts = false;
                _Workbook xlsWorkbook = xlsxApp.Workbooks.Open(mailExcel);
                _Worksheet xlsWorksheet = xlsWorkbook.Sheets[1];
                int k = xlsWorksheet.UsedRange.Rows.Count - 2;
                setText_RichTextBox(richtextbox, "正在导入邮件发送列表信息", Brushes.Blue);
                for (int i = 0; i < k; i++)
                {
                    recepientNameArray.Add(System.Convert.ToString(xlsWorksheet.Cells[i + 3, 1].Value));
                    recepientmailAddressArray.Add(System.Convert.ToString(xlsWorksheet.Cells[i + 3, 2].Value));
                    msgSubjectArray.Add(System.Convert.ToString(xlsWorksheet.Cells[i + 3, 3].Value));
                    msgHtmlBodyArray.Add(System.Convert.ToString(xlsWorksheet.Cells[i + 3, 4].Value));
                    attachmentsName.Add(System.Convert.ToString(xlsWorksheet.Cells[i + 3, 5].Value));
                    fileFullnameArray.Add(System.Convert.ToString(xlsWorksheet.Cells[i + 3, 6].Value));
                }
                setText_RichTextBox(richtextbox, "数组生成完毕", Brushes.Blue);               
                setText_RichTextBox(richtextbox, "准备发送邮件", Brushes.Blue);


                for (int i = 0; i < k; i++)
                {
                    jmail.Message msg = new jmail.Message();
                    msg.MailServerUserName = mailUsername;
                    msg.MailServerPassWord = myPassword;
                    msg.From = mailAddress;
                    msg.FromName = mailUsername;
                    msg.Encoding = "base64";
                    msg.Charset = "gb2312";
                    msg.Logging = true;
                    msg.Silent = true;
                    msg.Recipients.Clear();
                    msg.Attachments.Clear();
                    msg.AddRecipient((String)recepientmailAddressArray[i], (String)recepientNameArray[i]);                   
                    msg.Subject = System.Convert.ToString(msgSubjectArray[i]);

                    if (msgSubjectArray[i] == null)
                    {
                        msg.Subject = System.Convert.ToString(attachmentsName[i]);
                    }

                    msg.HTMLBody = System.Convert.ToString(msgHtmlBodyArray[i]);
                    if (msg.HTMLBody == "")
                    {
                        msg.HTMLBody = System.Convert.ToString(attachmentsName[i]);
                    }

                    if (!((String)fileFullnameArray[i] == ""))
                    {
                        msg.AddAttachment((String)fileFullnameArray[i]);

                    }
                    FileInfo fileInfo = new FileInfo((String)fileFullnameArray[i]);
                    
                    if (msg.Send(smtpServer))
                    {
                        setText_RichTextBox(richtextbox, (String)fileFullnameArray[i] + "：附件大小：" + ((fileInfo.Length) / 1024) + "KB" + "   发送成功√\n", Brushes.Green, NOTCLEAR);
                        sucessful++;
                        xlsWorksheet.Cells[i + 3, 7].Interior.ColorIndex = 4;
                        xlsWorksheet.Cells[i + 3, 7].value = "发送成功√";

                    }
                    else
                    {
                        setText_RichTextBox(richtextbox, (String)fileFullnameArray[i] + "：附件大小：" + ((fileInfo.Length) / 1024) + "KB" + "   发送失败×\n"  + msg.ErrorMessage, Brushes.Red, NOTCLEAR);
                        failure++;
                        xlsWorksheet.Cells[i + 3, 7].Interior.ColorIndex = 3;
                        xlsWorksheet.Cells[i + 3, 7].value = "发送失败×" + msg.Log;
                    }
                    msg.Clear();
                    msg.Close();
                    progressbar_display(i+1,k);

                }
                xlsWorksheet.Cells[2, 7].value = "发送状态";
                xlsWorkbook.Save();
                xlsWorkbook.Close();
                setText_RichTextBox(richtextbox, "邮件程序运行结束，已经在表格中用颜色标示是否发送成功！\n" + "绿色：发送成功   红色：发送失败\n" + "发送邮件共" + k + "封,  " + "成功" + sucessful + "封,   " + "失败" + failure + "封", Brushes.Blue, NOTCLEAR);
                if (checkBox_viewSendResultIsChecked)
                {
                    System.Diagnostics.Process.Start(mailExcel);
                }

            }

            catch (Exception error)
            {
                setText_RichTextBox(richtextbox, "发生异常，信息如下：\nMessage:" + error.Message + "\nStackTrace" + error.StackTrace, Brushes.Red, CLEAR);
            }
        }



       
        //////////////////非核心代码：以下方法用于设置RichTextBox显示的text内容
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
        public void setText_RichTextBox(System.Windows.Controls.RichTextBox richTextBox, string text, SolidColorBrush color)
        {
            System.Windows.Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
            {
                richTextBox.Document.Blocks.Clear();
                richTextBox.Foreground = color;
                richTextBox.AppendText(text);
            });

        }

    }
}