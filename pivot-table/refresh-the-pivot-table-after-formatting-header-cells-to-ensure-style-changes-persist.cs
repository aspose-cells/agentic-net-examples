// Title: Refresh Aspose.Cells Pivot Table After Header Formatting in C# (.NET)
// Description: Demonstrates how to load a workbook, enable PreserveFormatting, apply a bold white‑on‑dark‑blue style to a pivot table header using PivotTable.Format, refresh the pivot with Worksheet.RefreshPivotTables, and save the updated file. Ensures custom header styles persist after a data refresh.
// Keywords: Aspose.Cells | C# pivot table refresh | PreserveFormatting property | PivotTable.Format header style | Worksheet.RefreshPivotTables | custom pivot header styling | .NET Excel pivot table | Aspose.Cells example
// Common Searches: how to keep pivot table header style after refresh Aspose.Cells | Aspose.Cells PreserveFormatting example C# | refresh pivot tables without losing formatting | apply custom style to pivot table header Aspose.Cells | Worksheet.RefreshPivotTables usage
// Developer Intent: Refresh a pivot table while preserving any custom header formatting applied through Aspose.Cells.
// Use Cases: Apply a bold white font on a dark blue background to a pivot table header and ensure the style survives a refresh. | Set PreserveFormatting = true before calling RefreshPivotTables to retain all custom pivot styles. | Load an existing workbook, verify the presence of a pivot table, style its header rows, refresh the pivot, and save the result.
// AI Prompts: Generate C# code that formats a pivot table header with a custom style and refreshes the pivot while preserving the formatting using Aspose.Cells. | Explain the interaction between PreserveFormatting and Worksheet.RefreshPivotTables in Aspose.Cells for .NET. | Create a snippet that checks for pivot tables, applies a custom style to multiple header rows, refreshes them, and saves the workbook.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to load a workbook, enable PreserveFormatting, apply a bold white‑on‑dark‑blue style to a pivot table header using PivotTable.Format, refresh the pivot with Worksheet.RefreshPivotTables, and save the updated file. Ensures custom header styles persist after a data refresh.
class RefreshPivotAfterHeaderFormatting
{
    static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Preserve formatting when the pivot table is refreshed
            pivotTable.PreserveFormatting = true;

            // Create a style for the header cells
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to a header cell (row 0, column 0 in pivot table coordinates)
            // Adjust the row/column indices as needed for your specific layout
            pivotTable.Format(0, 0, headerStyle);

            // Refresh the pivot table so that the formatting is retained
            worksheet.RefreshPivotTables();

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
