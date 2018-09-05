using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
namespace NetMail
{
    /// <summary>
    /// 发送邮件
    /// </summary>
    public class SendMai1Helper
    {
        #region 公共属性       
        /// <summary>
        /// 邮件主体,可设置邮件接收人,内容等对应参数(假如调用的是默认构造参数,此属性对应的相关参数必须设置)
        /// </summary>
        public MailMessage MMsg { get; }
        /// <summary>
        /// 发送主体,可设置发送邮件服务器,端口,密码等相关参数(假如调用的是默认构造参数,此属性对应的相关参数可以设置)
        /// </summary>
        public SmtpClient Smtp{get;}

        public delegate void SendCompletedBack(bool cancelled, Exception error);
        /// <summary>
        /// 邮件发送完成事件
        /// </summary>
        public event SendCompletedBack SendCompleted;
        #endregion

        #region 构造函数
        public SendMai1Helper()
        {
            MMsg = new MailMessage();
            Smtp = new SmtpClient();
            Smtp.SendCompleted += Smtp_SendCompleted;
        }

        private void Smtp_SendCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            SendCompleted(e.Cancelled, e.Error);
            foreach (var item in MMsg.Attachments)
                item.Dispose();
            Smtp.Dispose();
            MMsg.Dispose();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender">发件人邮箱地址(可随便写)</param>
        /// <param name="senderName">发件人姓名</param>
        /// <param name="receiver">收件人邮箱地址列表</param>
        /// <param name="subject">邮件主题</param>
        /// <param name="body">邮件内容</param>
        public SendMai1Helper(string sender, string senderName, List<string> receiver,string subject,string body)
        {
            MMsg = new MailMessage();
            MMsg.From = new MailAddress(sender, senderName);
            foreach (string item in receiver)
            {
                MMsg.To.Add(new MailAddress(item));
            }
            MMsg.Subject = subject;
            MMsg.Body = body;
            MMsg.SubjectEncoding = Encoding.UTF8;
            MMsg.BodyEncoding = Encoding.UTF8;
            MMsg.HeadersEncoding = Encoding.UTF8;
            MMsg.Priority = MailPriority.High;
            MMsg.IsBodyHtml = true;
            Smtp = new SmtpClient();
            Smtp.SendCompleted += Smtp_SendCompleted;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender">发件人邮箱地址(可随便写)</param>
        /// <param name="senderName">发件人姓名</param>
        /// <param name="receiver">收件人邮箱地址列表</param>
        /// <param name="subject">邮件主题</param>
        /// <param name="body">邮件内容</param>
        /// <param name="sendHost">SMTP服务器地址</param>
        /// <param name="port">SMTP服务器端口</param>
        /// <param name="sendMailaddr">与SMTP服务器商匹配的邮箱地址(一般是在某个网注册的邮箱)</param>
        /// <param name="password">与SMTP服务器商匹配的邮箱密码开启SMTP服务时的授权码(一般是在某个网注册的邮箱)</param>
        public SendMai1Helper(string sender, string senderName, List<string> receiver, string subject, string body, string sendHost, int port, string sendMailaddr, string password) : this(sender, senderName, receiver,subject,body)
        {
            Smtp.Host = sendHost;
            Smtp.Port = port;
            Smtp.Credentials = new NetworkCredential(sendMailaddr, password);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender">发件人邮箱地址(可随便写)</param>
        /// <param name="senderName">发件人姓名</param>
        /// <param name="receiver">收件人邮箱地址列表</param>
        /// <param name="subject">邮件主题</param>
        /// <param name="body">邮件内容</param>
        /// <param name="sendHost">SMTP服务器地址</param>
        /// <param name="port">SMTP服务器端口</param>
        /// <param name="enableSSL">SMTP服务器是否启用了SSL加密</param>
        /// <param name="sendMailaddr">与SMTP服务器商匹配的邮箱地址(一般是在某个网注册的邮箱)</param>
        /// <param name="password">与SMTP服务器商匹配的邮箱密码开启SMTP服务时的授权码(一般是在某个网注册的邮箱)</param>
        public SendMai1Helper(string sender, string senderName, List<string> receiver, string subject, string body, string sendHost, int port, bool enableSSL, string sendMailaddr, string password) : this(sender, senderName, receiver, subject, body, sendHost, port, sendMailaddr, password)
        {
            Smtp.EnableSsl = enableSSL;
        }
        #endregion

        #region 方法
        /// <summary>
        /// 增加附件
        /// </summary>
        /// <param name="fileList">文件全路径名称列表</param>
        public void AddAttachment(List<string> fileList)
        {
            MMsg.Attachments.Clear();
            foreach(string item in fileList)
            {
                MMsg.Attachments.Add(new Attachment(item, MediaTypeNames.Application.Octet));
            }
        }
        /// <summary>
        /// 设置SMTP服务器相关参数
        /// </summary>
        /// <param name="sendHost">SMTP服务器地址</param>
        /// <param name="port">SMTP服务器端口</param>
        /// <param name="sendMail">与SMTP服务器商匹配的邮箱地址(一般是在某个网注册的邮箱)</param>
        /// <param name="password">与SMTP服务器商匹配的邮箱密码开启SMTP服务时的授权码(一般是在某个网注册的邮箱)</param>
        /// <param name="enableSSL">SMTP服务器是否启用了SSL加密</param>
        public void SetSmtp(string sendHost,int port,string sendMail,string password,bool enableSSL=false)
        {
            Smtp.Host = sendHost;
            Smtp.Port = port;
            Smtp.Credentials = new NetworkCredential(sendMail, password);
            Smtp.EnableSsl = enableSSL;
            
        }
        /// <summary>
        /// 设置SMTP服务器相关参数
        /// </summary>
        /// <param name="sendHost">SMTP服务器地址</param>
        /// <param name="port">SMTP服务器端口</param>
        public void SetSmtp(string sendHost, int port)
        {
            Smtp.Host = sendHost;
            Smtp.Port = port;

        }
        /// <summary>
        /// 设置邮件信息类基本参数
        /// </summary>
        /// <param name="sender">发件人邮箱地址(可随便写)</param>
        /// <param name="senderName">发件人姓名</param>
        /// <param name="receiver">收件人邮箱地址列表</param>
        /// <param name="subject">邮件主题</param>
        /// <param name="body">邮件内容</param>
        /// <param name="encoding">邮件内容,主题,头内容编码;默认为UTF8</param>
        public void SetMailMessage(string sender, string senderName, List<string> receiver, string subject, string body,Encoding encoding=null)
        {
            MMsg.From = new MailAddress(sender, senderName);
            foreach (string item in receiver)
            {
                MMsg.To.Add(new MailAddress(item));
            }
            if (encoding == null)
                encoding = Encoding.UTF8;
            MMsg.Subject = subject;
            MMsg.Body = body;
            MMsg.SubjectEncoding = encoding;
            MMsg.BodyEncoding = encoding;
            MMsg.HeadersEncoding = encoding;
            MMsg.Priority = MailPriority.High;
            MMsg.IsBodyHtml = true;
        }
        /// <summary>
        /// 设置邮件信息类的抄送及暗抄送
        /// </summary>
        /// <param name="cc">抄送人邮件地址列表</param>
        /// <param name="bcc">暗抄送人邮件地址列表</param>
        public void SetCC(List<string> cc,List<string> bcc)
        {
            foreach (string item in cc)
                MMsg.CC.Add(new MailAddress(item));
            foreach (string item in bcc)
                MMsg.Bcc.Add(new MailAddress(item));
        }


        /// <summary>
        /// 同步发送邮件,同时返回错误字符串;为空时则正常发送成功
        /// </summary>
        /// <returns></returns>
        public string SendMail()
        {
            string result = string.Empty;
            try
            {
                Smtp.Send(MMsg);
            }
            catch (Exception e)
            {
                result = e.Message;
               
                
            }
            foreach (var item in MMsg.Attachments)
                item.Dispose();
            Smtp.Dispose();
            MMsg.Dispose();
            return result;
        }
        /// <summary>
        /// 异步发送邮件
        /// </summary>
        public void SendMailSync()
        {
            string result = string.Empty;
            try
            {
                Smtp.SendAsync(MMsg, null);
            }catch(Exception e)
            {
                foreach (var item in MMsg.Attachments)
                    item.Dispose();
                Smtp.Dispose();
                MMsg.Dispose();
                throw e;
            }
        }
        #endregion

    }
}
