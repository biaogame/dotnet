 public class MailHelper
    {
        public MailHelper(string authorName, string subject, string body, string to,Attachment attachment)
        {
            this.AuthorName = authorName;
            this.Tos = to;
            this.Subject = subject;
            this.Body = body;
            this.attachment = attachment;

        }

        public delegate int MethodDelegate(int x, int y);
        private readonly int smtpPort = Convert.ToInt32(SystemConfigure.EmailsmtpPort);
        readonly string SmtpServer = SystemConfigure.EmailSmtpServer;
        private readonly string UserName = SystemConfigure.EmailUserName;
        readonly string Pwd = SystemConfigure.EmailPwd;
        public string AuthorName { get; set; }
        // private readonly string AuthorName =  SystemConfigure.EmailAuthorName;
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Tos { get; set; }
        public bool EnableSsl { get; set; }
        public Attachment attachment { get; set; }
        MailMessage GetClient
        {
            get
            {

                if (string.IsNullOrEmpty(Tos)) return null;
                MailMessage mailMessage = new MailMessage();
                //多个接收者                
                foreach (string _str in Tos.Split(';'))
                {
                    mailMessage.To.Add(_str);
                }
                mailMessage.From = new System.Net.Mail.MailAddress(UserName, AuthorName);
                mailMessage.Subject = Subject;
                mailMessage.Body = Body;
                mailMessage.IsBodyHtml = true;
                mailMessage.BodyEncoding = System.Text.Encoding.UTF8;
                mailMessage.SubjectEncoding = System.Text.Encoding.UTF8;
                mailMessage.Priority = System.Net.Mail.MailPriority.High;
                mailMessage.Attachments.Add(attachment);
                return mailMessage;

            }
        }
        SmtpClient GetSmtpClient
        {
            get
            {
                return new SmtpClient
                {
                    EnableSsl = true,//qq
                    UseDefaultCredentials = false,//q
                    Credentials = new System.Net.NetworkCredential(UserName, Pwd),
                    Host = SmtpServer,
                    //DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network,
                    //Port = smtpPort,
                    //UseDefaultCredentials = false,
                    //EnableSsl = false,
                };
            }
        }
        //回调方法
        Action<MailResult> actionSendCompletedCallback = null;
        ///// <summary>
        ///// 使用异步发送邮件
        ///// </summary>
        ///// <param name="AuthorName">发送人昵称</param>
        ///// <param name="subject">主题</param>
        ///// <param name="body">内容</param>
        ///// <param name="to">接收者,以,分隔多个接收者</param>
        //// <param name="_actinCompletedCallback">邮件发送后的回调方法</param>
        ///// <returns></returns>
        public void SendAsync(string authorName, string subject, string body, string to, Attachment _attachment, Action<MailResult> _actinCompletedCallback)
        {
            try
            {
                if (string.IsNullOrEmpty(to)) return;
                AuthorName = authorName;
                Tos = to;
                Subject = subject;
                Body = body;
                this.attachment = _attachment;
                EnableSsl = false;
                SmtpClient smtpClient = GetSmtpClient;
                MailMessage mailMessage = GetClient;
                if (smtpClient == null || mailMessage == null) return;


                //发送邮件回调方法
                actionSendCompletedCallback = _actinCompletedCallback;
                smtpClient.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);
                try
                {
                    smtpClient.SendAsync(mailMessage, "true");//异步发送邮件,如果回调方法中参数不为"true"则表示发送失败
                }
                catch (Exception e)
                {
                    //throw new Exception(e.Message);
                    //PubMethod.CreateLogTxt("发送邮件失败：" + e.ToString());
                    LogHelper.Err("发送邮件出现错误", e.ToString());
                }
                finally
                {
                    smtpClient = null;
                    mailMessage = null;
                }
                
                System.Threading.Thread.Sleep(SystemConfigure.SendEmailDelayTime * 1000);//延迟执行
            }
            catch (Exception ex) { LogHelper.Err("发送邮件时出现错误", ex.ToString()); }
        }
        /// <summary>
        /// 异步操作完成后执行回调方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SendCompletedCallback(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
          
                //同一组件下不需要回调方法,直接在此写入日志即可
                //写入日志
                //return;
                if (actionSendCompletedCallback == null) return;
                string message = string.Empty;
                this.Subject = this.Subject == null ? "" : this.Subject;
                this.Tos = this.Tos == null ? "" : this.Tos;
                MailResult mRsult = new MailResult();
                mRsult.Subject = this.Subject;
                mRsult.Tos = this.Tos;
                if (e.Cancelled)
                {
                    message = "异步操作取消"; mRsult.IsSuccess = false;
                    message = (string.Format("发送取消：UserState:{0},Message:{1},邮件标题:{2},接收人:{3}", (string)e.UserState,e.Error==null?"": e.Error.ToString(),this.Subject, this.Tos));
                    mRsult.message = (string.Format("发送取消：UserState:{0},Message:{1},邮件标题:{2},接收人:{3}", (string)e.UserState, e.Error.ToString(), this.Subject, this.Tos));

                }
                else if (e.Error != null)
                {
                    mRsult.IsSuccess = false;

                    message = (string.Format("发送失败：UserState:{0},Message:{1},邮件标题:{2},接收人:{3}", (string)e.UserState, e.Error == null ? "" : e.Error.ToString(), this.Subject, this.Tos));
                    mRsult.message = message;
                }
                else
                {
                    mRsult.IsSuccess = true;
                    message = (string.Format("发送失败：UserState:{0},Message:{1},邮件标题:{2},接收人:{3}", (string)e.UserState,"", this.Subject, this.Tos));
                    mRsult.message = (string.Format("发送取消：UserState:{0},Message:{1},邮件标题:{2},接收人:{3}", (string)e.UserState,"", this.Subject, this.Tos));
                }
                //执行回调方法
                actionSendCompletedCallback(mRsult);
         
        }
    }