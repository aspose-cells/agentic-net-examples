using System;
using Aspose.Cells;
using System.IO;

class UnmergeRangeExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the previously merged range D4:F4 using Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range mergedRange = worksheet.Cells.CreateRange("D4", "F4");

            // Unmerge the range, restoring individual cells
            mergedRange.UnMerge();

            // Save the workbook
            string outputPath = "UnmergedOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}