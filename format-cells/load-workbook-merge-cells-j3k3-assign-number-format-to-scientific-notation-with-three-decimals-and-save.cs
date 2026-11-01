// Title: How to merge cells J3:K3 and set a three‑decimal scientific number format in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge range J3:K3 and apply custom scientific format 0.000E+00 | set scientific notation with three decimal places for merged cells in Aspose.Cells | load workbook, merge cells, and save with Aspose.Cells .NET example | apply custom number format to merged cells using Aspose.Cells C# | Aspose.Cells example for merging cells and formatting as scientific notation
// Tags: Aspose.Cells merge cells example | scientific notation custom format Aspose.Cells | C# style application to merged cells | Excel workbook save Aspose.Cells .NET | handling missing workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The C# program loads an existing Excel file (or creates a new workbook if the file is missing or cannot be opened), merges the range J3:K3 on the first worksheet, creates a style with the custom scientific format "0.000E+00" (three decimal places), applies this style to the merged cell, and saves the workbook to output.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException)
                {
                    // If the file is password‑protected or cannot be opened, fall back to a new workbook
                    workbook = new Workbook();
                }
            }
            else
            {
                workbook = new Workbook();
            }

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells J3:K3 (zero‑based indices: row 2, column 9, 1 row, 2 columns)
            sheet.Cells.Merge(2, 9, 1, 2);

            // Create a style with scientific notation (three decimal places)
            Style sciStyle = workbook.CreateStyle();
            sciStyle.Custom = "0.000E+00";

            // Apply the style to the merged cell (top‑left cell of the range)
            sheet.Cells[2, 9].SetStyle(sciStyle);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
