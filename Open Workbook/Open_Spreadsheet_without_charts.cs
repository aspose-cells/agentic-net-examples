using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Input spreadsheet that may contain charts
        string inputPath = "input.xlsx";

        // Output path for the spreadsheet without any charts
        string outputPath = "output_no_charts.xlsx";

        // Load the workbook from the file (uses the provided constructor)
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and remove every chart
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Remove charts by iterating backwards to avoid index issues
            for (int i = sheet.Charts.Count - 1; i >= 0; i--)
            {
                sheet.Charts.RemoveAt(i);
            }
        }

        // Save the modified workbook (uses the provided Save method)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}