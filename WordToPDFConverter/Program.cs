using Aspose.Words;
using System.Drawing;
using System.IO;

namespace WordToPDFConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            static void Main(string[] args)
            {
                string root = @"C:\Users\rflorentino\Downloads\Materiales relacionados\New_XML_Archivo_de_Conversion\WORD-FILE";
                string[] fileEntries = Directory.GetFiles(root);

                foreach (string fileEntry in fileEntries)
                {
                    var doc = new Document(fileEntry);
                    doc.Save(fileEntry.Replace("docx", "epub"));
                }
            }
        }
    }
}
