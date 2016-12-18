using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApplication1.myclass
{
    struct JmailStruct
    {
        string sender;
        string sendername;
        string mailserverusername;
        string mailserverpassword;
        string smtp;
        string recepient;
        string recepient_name;
        string subject;
        string body;
        string attachment;
        string send()
        {
            try { 
            jmail.Message msg = new jmail.Message();
            msg.MailServerUserName = this.mailserverusername;
            msg.MailServerPassWord = this.mailserverpassword;
            msg.From = this.sender;
            msg.FromName = this.sendername;
            msg.Encoding = "base64";
            msg.Charset = "gb2312";
            msg.Logging = true;
            msg.Silent = true;
            msg.Subject = this.subject;
            msg.Body = this.body;
            msg.AddAttachment(this.attachment);
            if (msg.Send(smtp))
            {
                return "发送成功√";

            }
            else
            {
                return "发送失败";
            }
            }
            catch (Exception e)
            {
                return e.Message;

            }


        }
    }


}
