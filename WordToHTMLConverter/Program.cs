using SautinSoft.Document;

namespace WordToHTMLConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ConvertFromFile();
        }

        /// <summary>
        /// Convert DOCX to HTML (file to file).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-docx-to-html-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromFile()
        {
            string inpFile = @"C:\Users\Kodake\Desktop\Template.docx";
            string outFile = @"C:\Users\Kodake\Desktop\Template.html";

            DocumentCore dc = DocumentCore.Load(inpFile);
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}
