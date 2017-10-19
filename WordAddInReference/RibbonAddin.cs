using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word.Controls;
using System.Windows.Forms;
using MSForms = Microsoft.Vbe.Interop.Forms;
using System.Drawing;
using Controls = Microsoft.Office.Tools.Word.Controls;
namespace WordAddInReference
{
    public partial class RibbonAddin
    {
        Document vstoDocument;
        Word.Document oDoc;
        private void RibbonAddin_Load(object sender, RibbonUIEventArgs e)
        {
            vstoDocument = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);

            Word.Bookmarks bookmarks = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks;
            MessageBox.Show(bookmarks.Count.ToString());
            foreach (Word.Bookmark bookm in bookmarks)
            {
                MessageBox.Show(bookm.Name);
                if (bookm.Name.StartsWith("CTButton"))
                {
                    addButton(vstoDocument, bookm.Range, bookm.Name);
                }
            }

            object rng = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.First.Range;

            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("test");
            oDoc = Globals.ThisAddIn.Application.ActiveDocument;

            Globals.ThisAddIn.Application.DocumentBeforePrint += Application_DocumentBeforePrint;
            vstoDocument.BeforePrint += new System.ComponentModel.CancelEventHandler(vstoDocument_BeforePrint);
        }

        private void Application_DocumentBeforePrint(Word.Document Doc, ref bool Cancel)
        {
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {


            Document vstoDocument = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
            Word.Document oDoc =  Globals.ThisAddIn.Application.ActiveDocument;
            
            string extName = DateTime.Now.Ticks.ToString();
            extName = extName.Substring(extName.Length - 3);

            string name = "CTButton" +  extName;

            if (Globals.Ribbons.RibbonAddin.button1 != null)
            {
                Word.Selection oSelection = Globals.ThisAddIn.Application.Selection;
                if (oSelection != null && oSelection.Range != null)
                {
                    try
                    {
                        // addButton(vstoDocument,oSelection.Range, name);

                        Word.Range range = oSelection.Range;
                        object obj = (object)range;
                        oDoc.Bookmarks.Add(name, ref obj);
                        // Globals.ThisAddIn.Application.ActiveDocument.Save();
                        MessageBox.Show("Added Button:" + vstoDocument.Bookmarks.Count);

                       //var ctl = oDoc.ContentControls.Add(Word.WdContentControlType.wdContentControlText, ref obj);

                       addButton(vstoDocument, oSelection.Range, name);
                       
                       Word.Bookmarks bookmarks = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks;
                       foreach (Word.Bookmark bookm in bookmarks)
                       {
                           MessageBox.Show(bookm.Name);
                           if (bookm.Name.StartsWith("CTButton"))
                           {
                              //addButton(vstoDocument, bookmarks[i].Range, bookmarks[i].Name);
                           }
                       }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Fail to add Button" + ex.ToString());
                    }
                }
            }
            else
            {
                vstoDocument.Controls.Remove(name);
            }
                   
        }

        void vstoDocument_BeforePrint(object sender, System.ComponentModel.CancelEventArgs e)
        {
            removeControls();
            //toogleControl(true);
        }

        void toogleControl(bool hide)
        {
            //throw new NotImplementedException();
            // Get all of the Windows Forms controls.
            foreach (object control in vstoDocument.Controls)
            {
                if (control is Controls.Button)
                {
                    var btn = (Controls.Button)control;
                    if (btn.Name.StartsWith("CTButton"))
                    {
                        if (hide) btn.Hide(); btn.Show();
                    }
                }
            }
        }

        void removeControls()
        {
            System.Collections.ArrayList controlsToRemove = new System.Collections.ArrayList();
            // Get all of the Windows Forms controls.
            foreach (object control in vstoDocument.Controls)
            {
                if (control is System.Windows.Forms.Control)
                {
                    controlsToRemove.Add(control);
                }
            }

            // Remove all of the Windows Forms controls from the document.
            foreach (object control in controlsToRemove)
            {
                vstoDocument.Controls.Remove(control);
            }
        }

        void addButton(Document doc, Word.Range range,string name)
        {
            var button = doc.Controls.AddButton(
                             range, 100, 30, name);
            button.BackColor = Color.Red;
            button.Text = name;
            button.Click += new EventHandler(button_Click);           
            //button.Show();
        }

        void button_Click(object sender, EventArgs e)
        {
            var btn = (Microsoft.Office.Tools.Word.Controls.Button)(sender);
            MessageBox.Show(btn.Text);           
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
           
        }
    }
}
