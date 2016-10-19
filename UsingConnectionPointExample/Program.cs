using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Word;
using WordApp = Microsoft.Office.Interop.Word.Application;

namespace UsingConnectionPointExample
{
    internal class Program
    {
        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.None)]
        public class WordAppEventsSink : ApplicationEvents4
        {
            /// <summary>
            /// DocumentOpen event handler
            /// </summary>
            /// <param name="doc">opened document</param>
            public void DocumentOpen(Document doc)
            {
                Console.WriteLine("\t Document {0} opened", doc.FullName);
            }

            #region All another events sink methods
            public void DocumentBeforeClose(Document Doc, ref bool Cancel)
            {
                // do nothing
            }

            public void DocumentBeforePrint(Document Doc, ref bool Cancel)
            {
                // do nothing
            }

            public void DocumentBeforeSave(Document Doc, ref bool SaveAsUI, ref bool Cancel)
            {
                // do nothing
            }

            public void DocumentChange()
            {
                // do nothing
            }

            public void DocumentSync(Document Doc, Microsoft.Office.Core.MsoSyncEventType SyncEventType)
            {
                // do nothing
            }

            public void EPostageInsert(Document Doc)
            {
                // do nothing
            }

            public void EPostageInsertEx(Document Doc, int cpDeliveryAddrStart, int cpDeliveryAddrEnd, int cpReturnAddrStart, int cpReturnAddrEnd, int xaWidth, int yaHeight, string bstrPrinterName, string bstrPaperFeed, bool fPrint, ref bool fCancel)
            {
                // do nothing
            }

            public void EPostagePropertyDialog(Document Doc)
            {
                // do nothing
            }

            public void MailMergeAfterMerge(Document Doc, Document DocResult)
            {
                // do nothing
            }

            public void MailMergeAfterRecordMerge(Document Doc)
            {
                // do nothing
            }

            public void MailMergeBeforeMerge(Document Doc, int StartRecord, int EndRecord, ref bool Cancel)
            {
                // do nothing
            }

            public void MailMergeBeforeRecordMerge(Document Doc, ref bool Cancel)
            {
                // do nothing
            }

            public void MailMergeDataSourceLoad(Document Doc)
            {
                // do nothing
            }

            public void MailMergeDataSourceValidate(Document Doc, ref bool Handled)
            {
                // do nothing
            }

            public void MailMergeDataSourceValidate2(Document Doc, ref bool Handled)
            {
                // do nothing
            }

            public void MailMergeWizardSendToCustom(Document Doc)
            {
                // do nothing
            }

            public void MailMergeWizardStateChange(Document Doc, ref int FromState, ref int ToState, ref bool Handled)
            {
                // do nothing
            }

            public void NewDocument(Document Doc)
            {
                // do nothing
            }

            public void ProtectedViewWindowActivate(ProtectedViewWindow PvWindow)
            {
                // do nothing
            }

            public void ProtectedViewWindowBeforeClose(ProtectedViewWindow PvWindow, int CloseReason, ref bool Cancel)
            {
                // do nothing
            }

            public void ProtectedViewWindowBeforeEdit(ProtectedViewWindow PvWindow, ref bool Cancel)
            {
                // do nothing
            }

            public void ProtectedViewWindowDeactivate(ProtectedViewWindow PvWindow)
            {
                // do nothing
            }

            public void ProtectedViewWindowOpen(ProtectedViewWindow PvWindow)
            {
                // do nothing
            }

            public void ProtectedViewWindowSize(ProtectedViewWindow PvWindow)
            {
                // do nothing
            }

            public void Quit()
            {
                // do nothing
            }

            public void Startup()
            {
                // do nothing
            }

            public void WindowActivate(Document Doc, Window Wn)
            {
                // do nothing
            }

            public void WindowBeforeDoubleClick(Selection Sel, ref bool Cancel)
            {
                // do nothing
            }

            public void WindowBeforeRightClick(Selection Sel, ref bool Cancel)
            {
                // do nothing
            }

            public void WindowDeactivate(Document Doc, Window Wn)
            {
                // do nothing
            }

            public void WindowSelectionChange(Selection Sel)
            {
                // do nothing
            }

            public void WindowSize(Document Doc, Window Wn)
            {
                // do nothing
            }

            public void XMLSelectionChange(Selection Sel, XMLNode OldXMLNode, XMLNode NewXMLNode, ref int Reason)
            {
                // do nothing
            }

            public void XMLValidationError(XMLNode XMLNode)
            {
                // do nothing
            }
            #endregion
        }

        private static void Main(string[] args)
        {
            Console.WriteLine("Monitoring all opened documents by Microsoft Word.");
            Console.WriteLine("(Using classic COM connection point)");
            Console.WriteLine();

            // Create Microsoft.Office.Interop.Word.Application:
            var wordApp = new WordApp();
            
            var sink = new WordAppEventsSink(); // create event sink object
            var connectionPointContainer = (IConnectionPointContainer)wordApp; // use COM like IConnectionPointContainer
            Guid guid = typeof(ApplicationEvents4).GUID; // got guid of Microsoft Word Object Library

            // getting connection point object:
            IConnectionPoint connectionPoint;
            connectionPointContainer.FindConnectionPoint(ref guid, out connectionPoint);

            // subscribe to all events through our WordAppEventsSink object:
            int cookie; 
            connectionPoint.Advise(sink, out cookie);
            
            // now we are listening...
            Console.WriteLine("Press <Enter> for exit.");
            Console.WriteLine();
            Console.ReadKey();

            //after all we use cookie for unsubscribe
            connectionPoint.Unadvise(cookie);
            
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); // forcing a quick release of COM object
        }

    }
}
