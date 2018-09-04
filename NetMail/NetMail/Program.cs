using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace NetMail
{
    class Program
    {
        static void Main(string[] args)
        {
            #region 发送邮件
            string receiver = "329986833@qq.com";
            string smtp = ConfigurationManager.AppSettings["MAIL_SMTP_HOST"];//smtp服务器
            int port = 25;//smtp服务器端口
            try
            {
                port = int.Parse(ConfigurationManager.AppSettings["MAIL_SMTP_PORT"]);
            }
            catch { }

            string smtp_mail_addr = ConfigurationManager.AppSettings["MAIL_ADDR"];//对应SMTP服务商的邮箱地址(以什么邮箱发送邮件)
            string smtp_mail_pass = ConfigurationManager.AppSettings["MAIL_PASS"];//对应SMTP服务商的邮箱口令或授权码
            bool enableSsl = false;//smtp服务器是否启用SSL加密
            try
            {
                enableSsl = Boolean.Parse(ConfigurationManager.AppSettings["MAIL_SSL"]);
            }
            catch { }

            string subject = "学生信息表"; string body = "学生信息表已发送,请查收附件";

            SendMai1Helper mail = new SendMai1Helper(smtp_mail_addr, "悦动网络科技有限公司", new List<string> { receiver }, subject, body);
            //mail.AddAttachment(new List<string> { filePath });
            mail.SetSmtp(smtp, port, smtp_mail_addr, smtp_mail_pass, true);
            mail.SendMail();//同步发送
            #endregion
        }
    }
}
