using System;
using Aspose.Cells;

class HideSecondWorksheet
{
    static void Main()
    {
        // Load the existing workbook from disk
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Hide the second worksheet (index 1) if it exists
        if (workbook.Worksheets.Count > 1)
        {
            // Set the worksheet's visibility to false
            workbook.Worksheets[1].IsVisible = false;
            // Alternatively, you could use:
            // workbook.Worksheets[1].VisibilityType = VisibilityType.Hidden;
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}