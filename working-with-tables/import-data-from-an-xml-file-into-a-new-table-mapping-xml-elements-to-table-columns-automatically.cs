using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Import the XML file into the first worksheet.
        // The XML elements are automatically mapped to table columns.
        // Parameters: XML file path, destination sheet name, start row (0‑based), start column (0‑based)
        workbook.ImportXml("data.xml", "Sheet1", 0, 0);

        // Save the workbook with the imported data
        workbook.Save("output.xlsx");
    }
}