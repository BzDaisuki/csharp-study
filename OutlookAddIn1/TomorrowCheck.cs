using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using static System.Environment;

namespace OutlookAddIn1
{
    public partial class TomorrowCheck
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        /// <summary>
        /// 金曜日のメール送信時に「明日」と書いてあったら警告を表示する
        /// </summary>
        /// <param name="Item"></param>
        /// <param name="Cancel"></param>
        public void Application_ItemSend(object Item, ref bool Cancel)
        {
            Outlook.MailItem mail = Item as Outlook.MailItem;
            bool isFriday = (int)DateTime.Today.DayOfWeek == 5 ;

#if DEBUG
            isFriday = true;
#endif
            if(isFriday && mail.Body.IndexOf("明日") > -1)
            { 
            //土曜出勤を自らほのめかしてはいけない
            DialogResult result = MessageBox.Show
                ($"ちょい待って！{NewLine}キミの「明日」、ひょっとして「来週」じゃない？{NewLine} ",
                "このまま送る？",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Exclamation,
                MessageBoxDefaultButton.Button2);

            if (result == DialogResult.No) Cancel = true;
            }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //注: Outlook はこのイベントを発行しなくなりました。Outlook が
            //    を Outlook のシャットダウン時に実行する必要があります。https://go.microsoft.com/fwlink/?LinkId=506785 をご覧ください
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
