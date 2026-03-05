using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be reviewed
        string excelPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(excelPath);

        // Simple summary of the workbook
        Console.WriteLine("Spreadsheet Summary:");
        Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int rows = sheet.Cells.MaxDataRow + 1;
            int cols = sheet.Cells.MaxDataColumn + 1;
            Console.WriteLine($"Sheet \"{sheet.Name}\": {rows} rows, {cols} columns");
        }
    }
}