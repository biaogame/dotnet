
class Program
    {
        static void Main(string[] args)
        {
            Attachment attachment = new Attachment(ContectZongbuTablefilePath, MediaTypeNames.Application.Octet);
           new MailHelper("AuthorName", "Subject", "body", "to", attachment).SendAsync("AuthorName", "Subject", "body", "to", attachment, emailCompleted);
        }

        /// <summary>
        /// 邮件发送后的回调方法
        /// </summary>
        /// <param name="message"></param>
        void emailCompleted(MailResult message)
        {
            if (message.IsSuccess)
            {
                LogHelper.Info("发送邮件完毕", message.message);
            }
            else
            {
                LogHelper.Err("发送邮件错误", message.message);
            }
            Console.WriteLine("发送邮件完毕:" + message.message);
        }

      
    }




