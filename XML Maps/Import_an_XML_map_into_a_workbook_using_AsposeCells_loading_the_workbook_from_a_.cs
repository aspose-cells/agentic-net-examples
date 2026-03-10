using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (workbook-load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Path to the XML file to be imported
        string xmlFilePath = "data.xml";

        // Ensure the destination worksheet exists; use the first sheet if "Sheet1" is absent
        Worksheet sheet;
        try
        {
            sheet = workbook.Worksheets["Sheet1"];
        }
        catch
        {
            sheet = workbook.Worksheets[0];
        }

        // Import the XML data into the worksheet starting at cell A1 (row 0, column 0)
        workbook.ImportXml(xmlFilePath, sheet.Name, 0, 0);

        // Save the updated workbook (workbook-save rule)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}