using System;
using Aspose.Cells;

class LoadExcel95Demo
{
    static void Main()
    {
        // Path to the Excel 95/5.0 workbook file (xls format)
        string filePath = "sample.xls";

        // Load the workbook using the constructor that accepts a file path
        Workbook workbook = new Workbook(filePath);

        // Iterate through all worksheets to confirm they are initialized
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}, Cells count: {sheet.Cells.Count}");
        }

        // Example manipulation: write a value to cell A1 of the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.Cells["A1"].PutValue("Loaded");

        // Optionally save the modified workbook to a new file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}