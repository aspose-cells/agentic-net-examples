// Title: Sync the checked state of multiple CheckBox shapes with worksheet cells in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that iterates over rows 2‑10 in column C of an Excel workbook, adds a CheckBox shape for each row with Aspose.Cells, and links the CheckBox's LinkedCell property to the corresponding cell. | Update the program so that only CheckBox shapes linked to cells within C2:C10 are removed, then recreate them with new positions while preserving any checkboxes outside that range, using Aspose.Cells. | Create a reusable C# method that takes startRow, endRow, columnIndex, and optional size parameters, and batch‑creates CheckBox shapes linked to the specified cell range with Aspose.Cells.
// Common Searches: aspnet cells c# add checkboxes linked to cells programmatically | how to delete and recreate checkboxes in a specific Excel range using Aspose.Cells | batch create linked checkbox shapes for rows in an Excel worksheet with Aspose.Cells .NET | synchronize checkbox visual state with boolean values in Excel using Aspose.Cells | remove existing checkboxes in column C and add new ones with Aspose.Cells C#
// Tags: Aspose.Cells linked checkbox shape creation | Aspose.Cells remove checkboxes by linked cell range | Aspose.Cells sync checkbox visual state with cell value | Aspose.Cells C# batch checkbox generation | Aspose.Cells worksheet checkbox management

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads DataTable.xlsx, deletes any CheckBox shapes linked to cells C2:C10, adds new CheckBox shapes for each row in that range, links each shape to its corresponding cell via the LinkedCell property, and saves the result as DataTable_Synced.xlsx.
class CheckBoxSync
{
    static void Main()
    {
        const string inputFile = "DataTable.xlsx";
        const string outputFile = "DataTable_Synced.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFile);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that contains the boolean values to sync with checkboxes
            int startRow = 1;   // zero‑based index (row 2 in Excel)
            int endRow   = 9;   // row 10
            int colIndex = 2;   // column C (zero‑based)

            // Remove any existing checkboxes that are linked to cells within the target range
            for (int i = sheet.CheckBoxes.Count - 1; i >= 0; i--)
            {
                CheckBox cb = sheet.CheckBoxes[i];
                if (string.IsNullOrEmpty(cb.LinkedCell))
                    continue;

                CellsHelper.CellNameToIndex(cb.LinkedCell, out int linkedRow, out int linkedCol);
                if (linkedCol == colIndex && linkedRow >= startRow && linkedRow <= endRow)
                {
                    sheet.CheckBoxes.RemoveAt(i);
                }
            }

            // Create a checkbox for each row and link it to the corresponding cell
            for (int row = startRow; row <= endRow; row++)
            {
                // Simple positioning (in pixels)
                int left   = 5;
                int top    = row * 20 + 5;
                int width  = 20;
                int height = 20;

                // Add the checkbox and obtain its object
                int cbIndex = sheet.CheckBoxes.Add(left, top, width, height);
                CheckBox checkBox = sheet.CheckBoxes[cbIndex];

                // Assign a name (optional)
                checkBox.Name = $"chk_{row + 1}";

                // Link the checkbox to the cell (A1 style address)
                string linkedCellAddress = CellsHelper.CellIndexToName(colIndex, row);
                checkBox.LinkedCell = linkedCellAddress;

                // The visual state of the checkbox is automatically reflected by the linked cell.
                // No need to set a separate Checked property (not available in this API version).
            }

            // Save the workbook with the synchronized states
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
