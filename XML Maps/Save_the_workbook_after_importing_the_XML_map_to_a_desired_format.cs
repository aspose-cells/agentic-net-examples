using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook(); // workbook-create

        // Import XML data into the first worksheet starting at cell A1
        // Ensure that "data.xml" exists in the application directory
        string xmlPath = "data.xml";
        string sheetName = "Sheet1";
        int startRow = 0;   // Row index (0‑based)
        int startColumn = 0; // Column index (0‑based)

        workbook.ImportXml(xmlPath, sheetName, startRow, startColumn); // ImportXml

        // Save the workbook in the desired format (XLSX in this example)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx); // workbook-save

        Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
    }
}