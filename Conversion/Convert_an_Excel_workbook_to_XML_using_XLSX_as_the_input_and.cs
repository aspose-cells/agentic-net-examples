using System;
using Aspose.Cells;

namespace AsposeCellsXmlConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the resulting XML file
            string outputPath = "output.xml";

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(sourcePath);

            // Create XML save options (default settings)
            XmlSaveOptions xmlSaveOptions = new XmlSaveOptions();

            // Save the workbook as an XML file using the save options
            workbook.Save(outputPath, xmlSaveOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been successfully converted to XML at '{outputPath}'.");
        }
    }
}