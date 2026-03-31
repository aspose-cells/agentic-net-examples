using System;
using Aspose.Cells;

namespace RibbonXmlPreserveDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook that may contain Ribbon XML
            string sourcePath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // Retrieve the existing Ribbon XML (if any)
            string existingRibbonXml = workbook.RibbonXml;

            // (Optional) Perform any workbook modifications here
            // For demonstration, we will just add a simple value to the first sheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Demo");

            // Restore the Ribbon XML to ensure it is preserved
            workbook.RibbonXml = existingRibbonXml;

            // Save the workbook to a new file, preserving the Ribbon UI definition
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}' with Ribbon XML preserved.");
        }
    }
}