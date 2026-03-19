using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Paths for the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the existing workbook from disk (Workbook(string) constructor)
        Workbook workbook = new Workbook(inputPath);

        // Perform a simple operation: write a value to cell A1 of the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Processed");

        // Save the modified workbook back to disk (Save(string) method)
        workbook.Save(outputPath);

        // Example of saving the same workbook in a different format (Save(string, SaveFormat) method)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}