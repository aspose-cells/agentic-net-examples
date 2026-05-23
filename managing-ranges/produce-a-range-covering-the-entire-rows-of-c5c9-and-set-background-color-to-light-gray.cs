using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a range that covers column C (index 2) rows 5‑9 (zero‑based rows 4‑8)
            // Parameters: startRow, startColumn, totalRows, totalColumns
            Aspose.Cells.Range cRange = cells.CreateRange(4, 2, 5, 1);

            // Expand the range to the entire rows that contain the cells in C5:C9
            Aspose.Cells.Range entireRows = cRange.EntireRow;

            // Define a style with a solid light‑gray background
            Style grayStyle = workbook.CreateStyle();
            grayStyle.Pattern = BackgroundType.Solid;
            grayStyle.ForegroundColor = Color.LightGray;

            // Apply the style to the entire rows range
            entireRows.SetStyle(grayStyle);

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}