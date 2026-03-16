using System;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    public class ExcelToXmlConverter
    {
        public static void Main()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the output XML file
            string outputPath = "output.xml";

            // Load the Excel workbook (create/load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Create XML save options (optional configuration)
            XmlSaveOptions saveOptions = new XmlSaveOptions();
            // Export sheet name as the XML element name
            saveOptions.SheetNameAsElementName = true;

            // Save the workbook as XML using the save options (save rule)
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been converted to XML at '{outputPath}'.");
        }
    }
}