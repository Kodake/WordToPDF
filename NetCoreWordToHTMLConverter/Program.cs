using DocXToPdfConverter;
using System;
using System.IO;
using System.Reflection;

namespace NetCoreWordToHTMLConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            /*
             * Your TODO:
             * 1. Enter your LibreOffice path below
             * 2. Have an input file with placeholders ready, in docx or HTML. Create your own placeholders (from the class Placeholders)
             * 3. Create an object for "ReportGenerator"
             * 4. Execute the method "Convert" on the object ReportGenerator.
             * Possible conversions: from HTML or from DOCX to PDF, HTML, DOCX
             *
             */

            //Enter the location of your LibreOffice soffice.exe below, full path with "soffice.exe" at the end
            //or anything you have in Linux...

            string locationOfLibreOfficeSoffice =
                @"C:\Users\Kodake\Downloads\LibreOfficePortable\LibreOfficePortable.exe";


            //This is only to get this example to work (find the word docx and the html file, which were
            //shipped with this).
            string executableLocation = Path.GetDirectoryName(
                Assembly.GetExecutingAssembly().Location);

            //Here are the 2 test files as input. They contain placeholders
            string docxLocation = Path.Combine(executableLocation, @"C:\Users\Kodake\Desktop\Template.docx");
            //string htmlLocation = Path.Combine(executableLocation, @"C:\Users\Kodake\Desktop\Template.html");

            //Most important: give the full path to the soffice.exe file including soffice.exe.
            //Don't know how that would be named on Linux...
            var test = new ReportGenerator(locationOfLibreOfficeSoffice);

            //Convert from DOCX to HTML
            test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(docxLocation), "Test-Template-out.html"), null);

            //Convert from DOCX to PDF
        }
    }
}
