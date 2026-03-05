using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the default worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Optional: import XML data if the file exists
        string xmlFile = "data.xml";
        if (File.Exists(xmlFile))
        {
            // Import XML into the worksheet starting at cell A1 (lifecycle rule: import)
            workbook.ImportXml(xmlFile, sheet.Name, 0, 0);
        }

        // Save the workbook to disk as an XLSX file (lifecycle rule: save)
        workbook.Save("output.xlsx");
    }
}