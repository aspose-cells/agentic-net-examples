using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Load the workbook from the file (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve worksheet name and the value of cell A1
        string sheetName = sheet.Name;
        string a1Value = sheet.Cells["A1"].StringValue;

        // Create an anonymous object containing the extracted information
        var workbookInfo = new
        {
            SheetName = sheetName,
            CellA1 = a1Value
        };

        // Display the anonymous object's properties
        Console.WriteLine($"Sheet Name: {workbookInfo.SheetName}");
        Console.WriteLine($"Cell A1 Value: {workbookInfo.CellA1}");

        // Save a copy of the workbook (uses the Workbook.Save method)
        workbook.Save("copy.xlsx", SaveFormat.Xlsx);
    }
}