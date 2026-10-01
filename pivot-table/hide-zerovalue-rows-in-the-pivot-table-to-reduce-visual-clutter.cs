// Title: Hide rows that contain only zero values in an Excel pivot table using Aspose.Cells for C#
// AI Prompts: Write a C# program with Aspose.Cells that refreshes a pivot table, calculates its data, and hides every worksheet row where all cells are zero. | Generate .NET code that scans the used range after pivot calculation and sets Cells.Rows[row].IsHidden = true for rows consisting solely of numeric zeros.
// Common Searches: Aspose.Cells C# hide rows with all zero values after pivot table refresh | C# code to programmatically hide zero-value rows in an Excel pivot table | How to remove empty rows from a pivot table using Aspose.Cells .NET | Iterate worksheet used range and hide rows where every cell equals 0 in C#
// Tags: hide zero rows Aspose.Cells | refresh calculate pivot Aspose.Cells | row hiding based on all-zero values C# | iterate worksheet used range Aspose.Cells | programmatic row concealment Excel C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, accesses the first worksheet's pivot table, refreshes and calculates it, then iterates through the worksheet's used range. Any row where every cell contains a numeric zero (or is null/empty) is marked as hidden, and the workbook is saved with the modified layout.
class HideZeroValuePivotRows
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first pivot table on the sheet (if any)
            if (worksheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No pivot tables found on the first worksheet.");

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Refresh and calculate the pivot table to ensure data is up‑to‑date
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Determine the used range of the worksheet (covers the pivot table data area)
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            // Iterate through each row in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                bool allZero = true;

                // Check each cell in the current row
                for (int col = 0; col <= maxCol; col++)
                {
                    object value = worksheet.Cells[row, col].Value;

                    // Treat null or empty as zero as well
                    if (value != null && double.TryParse(value.ToString(), out double numericValue))
                    {
                        if (numericValue != 0)
                        {
                            allZero = false;
                            break;
                        }
                    }
                    else
                    {
                        // Non‑numeric values are considered non‑zero for hiding purposes
                        allZero = false;
                        break;
                    }
                }

                // Hide the entire row if all values are zero
                if (allZero)
                {
                    // Use the Cells.Rows collection to hide the row
                    worksheet.Cells.Rows[row].IsHidden = true;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
