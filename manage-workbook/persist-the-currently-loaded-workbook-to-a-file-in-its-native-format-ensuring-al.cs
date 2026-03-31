using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook from a file.
        // The constructor Workbook(string) opens the file and creates the workbook object.
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Perform any modifications you need.
        // Example: write a value to cell A1 of the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Modified");

        // Persist the workbook to disk in its native format.
        // Save(string) determines the format from the file extension.
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);

        // Release resources.
        workbook.Dispose();
    }
}