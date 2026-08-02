using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeImport
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel workbook
            string workbookPath = "ExistingWorkbook.xlsx";

            // Path to the XML file that contains named range definitions
            string xmlPath = "NamedRanges.xml";

            // Load the existing workbook
            Workbook wb = new Workbook(workbookPath);

            // Import the XML data (including named range definitions) into the first worksheet,
            // starting at cell A1 (row 0, column 0)
            wb.ImportXml(xmlPath, "Sheet1", 0, 0);

            // Optional: list all named ranges after import to verify they were added
            NameCollection names = wb.Worksheets.Names;
            Console.WriteLine($"Total named ranges after import: {names.Count}");
            foreach (Name name in names)
            {
                Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
            }

            // Save the updated workbook
            string outputPath = "WorkbookWithImportedNames.xlsx";
            wb.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}