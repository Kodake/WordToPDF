using GemBox.Document;
using System;

namespace GemBoxWordToPNGConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Load a Word file into the DocumentModel object.
            var document = DocumentModel.Load(@"C:\Users\Kodake\Desktop\Template.docx");

            // Create image save options.
            var imageOptions = new ImageSaveOptions(ImageSaveFormat.Png)
            {
                PageNumber = 0, // Select the first Word page.
                Width = 1240 // Set the image width and keep the aspect ratio.
            };

            // Save the DocumentModel object to a PNG file.
            document.Save(@"C:\Users\Kodake\Desktop\Template.png", imageOptions);
        }
    }
}
