// Title: Loop through pivot table header cells by display name and apply custom styling with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells iterate over pivot table header cells by field name | Apply background color to specific pivot table headers using Aspose.Cells .NET | Find and style pivot field headers by display name in an Excel workbook with Aspose.Cells | How to change font style of pivot table column headers programmatically in C# | Search pivot table area for header text and set cell style using Aspose.Cells
// Tags: pivot table header styling Aspose.Cells | loop through pivot table cells C# | set cell background color based on header text Aspose | modify pivot field header font Aspose.Cells | search pivot table range for string value .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, checks for a pivot table on the first worksheet, and defines target field names (Region, Category, Year). It retrieves the pivot table's overall cell range and loops through each cell to locate headers matching the target names (case‑insensitive). When a match is found, a light‑blue solid background and bold dark‑blue font are applied to the header cell. The modified workbook is saved to a new file, with error handling for missing input files and save failures.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Define the display names of the pivot fields whose headers we want to format
            string[] targetFields = { "Region", "Category", "Year" };

            // Get the overall area of the pivot table (including headers)
            CellArea tableArea = pivotTable.TableRange1;

            // Loop through each target field name
            foreach (string fieldName in targetFields)
            {
                bool found = false;

                // Scan the pivot table area for a cell whose value matches the field name
                for (int row = tableArea.StartRow; row <= tableArea.EndRow && !found; row++)
                {
                    for (int col = tableArea.StartColumn; col <= tableArea.EndColumn && !found; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsString &&
                            cell.StringValue.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Apply style to the header cell
                            Style style = cell.GetStyle();
                            style.ForegroundColor = Color.LightBlue;
                            style.Pattern = BackgroundType.Solid;
                            style.Font.IsBold = true;
                            style.Font.Color = Color.DarkBlue;
                            cell.SetStyle(style);
                            found = true;
                        }
                    }
                }

                if (!found)
                {
                    Console.WriteLine($"Header for field \"{fieldName}\" not found in the pivot table.");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
