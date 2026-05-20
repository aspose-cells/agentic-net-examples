using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlImportAndRecalc
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Path to the XML file to be imported
            string xmlPath = "data.xml";

            // Import XML data into the first worksheet starting at cell A1
            // This uses the provided ImportXml method (string, string, int, int)
            workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

            // Recalculate all formulas in the workbook so that any dependent
            // calculations reflect the newly imported data.
            // This uses the provided CalculateFormula method.
            workbook.CalculateFormula();

            // Optionally, save the workbook to verify the results
            string outputPath = "ResultAfterImport.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"XML imported and formulas recalculated. Workbook saved to '{outputPath}'.");
        }
    }
}