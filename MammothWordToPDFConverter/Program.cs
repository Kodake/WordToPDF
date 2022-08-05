using Mammoth;

namespace MammothWordToPDFConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var converter = new DocumentConverter();
            var result = converter.ConvertToHtml(@"C:\Users\Kodake\Desktop\Template.docx");
            var html = result.Value; // The generated HTML
            var warnings = result.Warnings; // 
        }
    }
}
