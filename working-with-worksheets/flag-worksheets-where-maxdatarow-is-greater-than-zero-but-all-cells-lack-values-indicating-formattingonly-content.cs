// Title: Identify and flag Excel worksheets that contain only formatting (no cell values) by adding a custom property and setting the tab color with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to mark worksheets that have formatting but no data | detect Excel sheets where MaxDataRow is greater than zero but cells are empty using Aspose.Cells | add a custom property and change tab color for formatting‑only worksheets in Aspose.Cells | find worksheets that contain only styles and no cell content with Aspose.Cells for .NET | programmatically flag formatting‑only worksheets in a .xlsx file using Aspose.Cells
// Tags: detect formatting‑only worksheets Aspose.Cells | worksheet custom property Aspose.Cells | set worksheet tab color Aspose.Cells | maxdatarow empty values check C# | scan cells for non‑empty values Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, iterates through each worksheet, and uses Cells.MaxDataRow/MaxDataColumn to define the data range. It scans the cells within that range for any non‑null, non‑empty values. If no actual values are found, it adds a custom property "FormattingOnly" set to "true" and changes the worksheet's tab color to LightGray, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the last row/column that contain any data or formatting
                int maxDataRow = cells.MaxDataRow;       // zero‑based index
                int maxDataColumn = cells.MaxDataColumn; // zero‑based index

                // Proceed only if there is at least one row with data/formatting
                if (maxDataRow >= 0 && maxDataColumn >= 0)
                {
                    bool hasActualValue = false;

                    // Scan cells up to the detected bounds
                    for (int row = 0; row <= maxDataRow && !hasActualValue; row++)
                    {
                        for (int col = 0; col <= maxDataColumn; col++)
                        {
                            Cell cell = cells[row, col];

                            // Check for a non‑null, non‑empty value
                            if (cell.Value != null && !(cell.Value is string s && string.IsNullOrEmpty(s)))
                            {
                                hasActualValue = true;
                                break;
                            }
                        }
                    }

                    // If no real values were found, mark the sheet as formatting‑only
                    if (!hasActualValue)
                    {
                        // CustomProperties expects string values
                        sheet.CustomProperties.Add("FormattingOnly", "true");

                        // Change tab color for visual identification
                        sheet.TabColor = Color.LightGray;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
