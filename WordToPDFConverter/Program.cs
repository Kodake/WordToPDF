using Aspose.Words;
using System.Drawing;

namespace WordToPDFConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var doc = new Document(@"C:\Users\Kodake\Desktop\Template.docx");

            TextWatermarkOptions options = new TextWatermarkOptions()
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Black,
                Layout = WatermarkLayout.Horizontal,
                IsSemitrasparent = false
            };

            doc.Watermark.SetText("Test Document Text", options);

            doc.Save(doc + "AddTextWatermark_out.docx");

            //doc.Save(@"C:\Users\Kodake\Desktop\Template.pdf");
        }
    }
}
