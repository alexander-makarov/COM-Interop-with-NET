using System;
using Microsoft.Office.Interop.Word;
using WordApp = Microsoft.Office.Interop.Word.Application;

namespace UsingDelegatesExample
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Monitoring all opened documents by Microsoft Word.");
            Console.WriteLine("(Using .NET delegates)");
            Console.WriteLine();

            // Create Microsoft.Office.Interop.Word.Application:
            var wordApp = new WordApp();
            // subscribe to DocumentOpen event:
            wordApp.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(wordApp_DocumentOpen);

            Console.WriteLine("Press <Enter> for exit.");
            Console.WriteLine();
            Console.ReadKey();

            // unsubscribe:
            wordApp.DocumentOpen -= wordApp_DocumentOpen;
        }

        /// <summary>
        /// DocumentOpen event handler
        /// </summary>
        /// <param name="doc">opened document</param>
        private static void wordApp_DocumentOpen(Document doc)
        {
            Console.WriteLine("\t Document {0} opened", doc.FullName);
        }
    }
}
