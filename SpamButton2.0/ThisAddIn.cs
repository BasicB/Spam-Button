/*
 *        Copyright (C) 2017  BasicB
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
  */





using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace SpamButton2._0
{
    public partial class ThisAddIn
    {

        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        Outlook.Inspectors inspectors;
        public Boolean attachments;
        public Boolean outsideCMSAddress;
        public String finalPassword = null;

        public string number = null;
        public string carrier = null;
        Office.CommandBarButton btn1;
        public static int Count = 0;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inspectors = this.Application.Inspectors;
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);
            Outlook.Explorer explorer = Application.ActiveExplorer();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon2();
        }

        public static void deleteSpam()
        {
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            Outlook.Selection selection = explorer.Selection;

            if (selection.Count > 0)   // Check that selection is not empty.
            {
                object selectedItem = selection[1];   // Index is one-based.
                Outlook.MailItem mailItem = selectedItem as Outlook.MailItem;


                if (mailItem != null)    // Check that selected item is a message.
                {
                    try
                    {

                        Outlook.MailItem tosend = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
                        tosend.Attachments.Add(mailItem);
                        tosend.To = "Example@gmail.com";//Set the To equal to whatever the email address is for the spam mailbox.;
                        tosend.Subject = mailItem.Subject;
                        DateTime now = DateTime.Now;
                        tosend.Body = now.ToString();
                        tosend.Save();
                        tosend.Send();
                        mailItem.Delete();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("{0} Exception caught: ", ex);
                    }
                }

            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

 

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
