using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMerge
{
    class MergeJob
    {
        private string attachment;
        private string recipient;
        private List<string> carbonCopy = new List<string>();

        public MergeJob(string[] columns)
        {
            attachment = columns[0];
            recipient = columns[1];
            for (int i = 2; i < columns.Count() - 1; i++) carbonCopy.Add(columns[i]);
        }

        public override string ToString()
        {
            List<string> parts = new List<string>();
            parts.Add(attachment);
            parts.Add(recipient);
            parts.Concat(carbonCopy);
            StringBuilder result = new StringBuilder();
            string separator = "";
            foreach (string part in parts)
            {
                result.Append(separator);
                separator = ",";
                result.Append('"' + part.Replace("\"", "\"\"") + '"');
            }
            return result.ToString();
        }

        internal void execute(string emailTemplate, string emailServer, string emailUsername, string emailPassword)
        {
            var message = MimeMessage.Load(File.OpenRead(emailTemplate));
            message.To.Clear();
            message.Cc.Clear();
            message.To.Add(new MailboxAddress("", recipient));
            message.Date = DateTimeOffset.Now;
            message.MessageId = MimeUtils.GenerateMessageId();
            foreach (string cc in this.carbonCopy) message.Cc.Add(new MailboxAddress("", cc));
            var newBody = new Multipart("mixed");
            newBody.Add(message.Body); 
            newBody.Add(new MimePart("application", "pdf")
            {
                Content = new MimeContent(File.OpenRead(this.attachment),
                                          ContentEncoding.Default),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = Path.GetFileName(this.attachment)
            });
            
            message.Body = newBody;
            using (var client = new SmtpClient())
            {
                client.Connect(emailServer, 465, SecureSocketOptions.SslOnConnect);
                client.Authenticate(emailUsername, emailPassword);
                client.Send(message);
                client.Disconnect(true);
            }
        }
    }
}
