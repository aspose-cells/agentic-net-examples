// Title: Apply a custom header style to a pivot table and save the workbook as XLSX using Aspose.Cells for .NET
// Common Searches: how to add a data field to an existing pivot table using Aspose.Cells C# | C# code to style the header row of a pivot table with Aspose.Cells | refresh a pivot table after modifying fields in Aspose.Cells .NET | save a modified workbook as XLSX after pivot table changes with Aspose.Cells | apply custom header formatting to pivot table cells in Aspose.Cells for .NET
// Tags: apply custom style to pivot table header Aspose.Cells | add data field to pivot table C# Aspose.Cells | refresh pivot table after modification Aspose.Cells | save workbook as XLSX Aspose.Cells .NET | pivot table styling workbook export Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing XLSX workbook, checks for a pivot table on the first worksheet, adds a data field, refreshes the pivot table, creates a light‑blue bold style, applies it to the pivot table's header row, and saves the updated workbook as a new XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Example modification: add a data field (e.g., column index 2)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

            // Refresh the pivot table to apply the changes
            pivotTable.RefreshData();

            // Create a custom style for the pivot table header
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.DarkBlue;

            // Determine the header row range of the pivot table
            int headerRow = pivotTable.RowRange.StartRow;
            int firstCol = pivotTable.ColumnRange.StartColumn;
            int lastCol = pivotTable.ColumnRange.EndColumn;

            // Apply the custom style to each cell in the header row
            for (int col = firstCol; col <= lastCol; col++)
            {
                Cell cell = worksheet.Cells[headerRow, col];
                StyleFlag flag = new StyleFlag { All = true };
                cell.SetStyle(headerStyle, flag);
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
