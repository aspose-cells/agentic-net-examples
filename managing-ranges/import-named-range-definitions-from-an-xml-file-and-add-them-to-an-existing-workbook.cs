using System;
using Aspose.Cells;

class ImportNamedRanges
{
    static void Main()
    {
        // Path to the existing workbook
        string workbookPath = "ExistingWorkbook.xlsx";

        // Path to the XML file that contains named range definitions
        string xmlPath = "NamedRanges.xml";

        // Load the existing workbook
        Workbook workbook = new Workbook(workbookPath);

        // Import the XML data into the first worksheet starting at cell A1 (row 0, column 0)
        // This operation also brings in any named range definitions present in the XML.
        workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

        // Display the imported named ranges for verification
        Console.WriteLine("Imported Named Ranges:");
        foreach (Name name in workbook.Worksheets.Names)
        {
            Console.WriteLine($"{name.Text} => {name.RefersTo}");
        }

        // Save the workbook with the newly added named ranges
        workbook.Save("UpdatedWorkbook.xlsx");
    }
}