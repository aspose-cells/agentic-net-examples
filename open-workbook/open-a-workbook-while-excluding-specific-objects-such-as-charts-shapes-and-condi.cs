using System;
using Aspose.Cells;

class LoadWorkbookExcludingObjects
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook normally
        Workbook workbook = new Workbook(sourcePath);

        // Remove charts, drawings (shapes), and conditional formatting from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Remove all charts
            sheet.Charts.Clear();

            // Remove all shapes (drawings)
            sheet.Shapes.Clear();

            // Remove all conditional formatting rules
            sheet.ConditionalFormattings.Clear();
        }

        // Save the workbook to verify that the unwanted objects were removed
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine("Workbook loaded and cleaned of charts, drawings, and conditional formatting.");
        Console.WriteLine("Saved to: " + outputPath);
    }
}