using System;
using Aspose.Cells;

class ImportXmlDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook wb = new Workbook();

        // Import XML data from a file into the first worksheet starting at cell A1 (row 0, column 0)
        wb.ImportXml("data.xml", "Sheet1", 0, 0);

        // Save the workbook to an Excel file
        wb.Save("output.xlsx");
    }
}