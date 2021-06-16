using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckOutlookEmailDemo {
	
    class Program {

        private static NameSpace nameSpace;
        private static ApplicationClass outlookApp = new ApplicationClass();

        /// <summary>
        /// 获取新邮件中的内容
        /// </summary>
        /// <param name="entry"></param>
        private static void AnalyzeNewItem(string entry) {
            var inbox = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            List<MailItem> allmails = new List<MailItem>();
            foreach(var item in inbox.Items) {
                if(item is MailItem) {
                    var mail = item as MailItem;
                    allmails.Add(mail);
                }
            }
			// 拿到邮件中最新收到的那个邮件
            var latest = allmails.Max(s => s.ReceivedTime);
            var latestMailItem = allmails.FirstOrDefault(s => s.ReceivedTime == latest);
            if(latestMailItem != null) {
                Console.WriteLine("Sujet : " + latestMailItem.Subject);
                Console.WriteLine("TO : " + latestMailItem.To);
                Console.WriteLine("SenderName : " + latestMailItem.SenderName);
                Console.WriteLine("ReceivedTime" + latestMailItem.ReceivedTime);
				
				// 重点是如何从邮件的body内容中进行分析 ????? 
                Console.WriteLine("Body : " + latestMailItem.Body);
				
				// 存储一个邮件的有效信息到指定的Excel ==> 是否需要备份，作为缓冲，何时Upload 
            }
        }

        /// <summary>
        /// 根据新的邮件的信息，完成自动化的操作 ......
        /// </summary>
        private static void outlookApp_NewMail() {
            Console.WriteLine("To do something ......");
        }

        /// <summary>
        /// 监听新邮件的事件处理器：获取信息
        /// </summary>
        /// <param name="EntryIDCollection"></param>
        private static void outlookApp_NewMailEx(string EntryIDCollection) {
            Console.WriteLine("A new message comes");
            AnalyzeNewItem(EntryIDCollection);
        }

        /// <summary>
        /// Start listening to the new email comes 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args) {
            Console.WriteLine("start to monitor new emails");
            nameSpace = outlookApp.GetNamespace("MAPI");
            outlookApp.NewMailEx += new ApplicationEvents_11_NewMailExEventHandler(outlookApp_NewMailEx);
            outlookApp.NewMail += new ApplicationEvents_11_NewMailEventHandler(outlookApp_NewMail);
            while(true) {

            }
        }

    }
}
