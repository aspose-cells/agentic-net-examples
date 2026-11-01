// Title: Identify and flag shape‑only worksheets (no data rows/columns) by coloring the tab and adding a note using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells find worksheets that contain only shapes and no data rows | how to mark shape‑only Excel sheets with a red tab using Aspose.Cells | set worksheet tab color programmatically for sheets that have drawings but no cell data in Aspose.Cells | add a warning label to cell A1 of shape‑only worksheets with Aspose.Cells .NET | detect empty data range with MaxDataRow -1 and MaxDataColumn -1 in Aspose.Cells
// Tags: detect shape‑only worksheets Aspose.Cells | set worksheet tab color programmatically Aspose.Cells | add cell note to flag shape‑only sheet Aspose.Cells | use MaxDataRow and MaxDataColumn to identify empty data Aspose.Cells | iterate worksheets and check Shapes collection Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The C# program loads an Excel workbook (or creates a new one if the file is missing), iterates through each worksheet, and flags sheets where the Shapes collection is non‑empty while Cells.MaxDataRow and Cells.MaxDataColumn are both -1. For those shape‑only worksheets it sets the tab color to red, writes "Shape‑only sheet" in cell A1 with bold red formatting, and saves the workbook to the specified output file.
class ShapeOnlySheetFlagger
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load workbook if the file exists; otherwise create a new empty workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with a default sheet
            }

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the sheet contains shapes but no data rows/columns
                bool hasShapes = sheet.Shapes.Count > 0;
                bool noDataRows = sheet.Cells.MaxDataRow == -1;
                bool noDataColumns = sheet.Cells.MaxDataColumn == -1;

                if (hasShapes && noDataRows && noDataColumns)
                {
                    // Flag the sheet by setting its tab color to red
                    sheet.TabColor = Color.Red;

                    // Add a note in cell A1 for visual confirmation
                    Cell flagCell = sheet.Cells["A1"];
                    flagCell.PutValue("Shape‑only sheet");

                    Style style = flagCell.GetStyle();
                    style.ForegroundColor = Color.Red;
                    style.Font.IsBold = true;
                    flagCell.SetStyle(style);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
