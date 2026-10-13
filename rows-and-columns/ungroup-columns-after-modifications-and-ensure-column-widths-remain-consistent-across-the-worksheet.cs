// Title: How to ungroup columns B‑E and set a uniform column width in an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that ungroups columns 2‑5 in a worksheet and then copies the width of column A to all columns. | Show a .NET example that groups columns B through E, removes the grouping, and enforces the same column width for every column in the first sheet.
// Common Searches: Aspose.Cells C# ungroup a range of columns after grouping | C# set identical column width for all columns in an Excel workbook using Aspose.Cells | preserve column width when ungrouping columns with Aspose.Cells .NET | how to copy column width from first column to all columns in Aspose.Cells | Aspose.Cells .NET example for grouping and then ungrouping columns
// Tags: Aspose.Cells UngroupColumns method C# | Aspose.Cells SetColumnWidth API .NET | Aspose.Cells column grouping and ungrouping example | Aspose.Cells uniform column width worksheet | Aspose.Cells preserve column width after ungrouping

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, groups columns B‑E, immediately ungroups them, copies the width of column A to every column in the first worksheet, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Group columns B (index 1) through E (index 4)
            // The second parameter is the total number of columns to group (4 columns)
            sheet.Cells.GroupColumns(1, 4, true);

            // Ungroup the previously grouped columns
            sheet.Cells.UngroupColumns(1, 4);

            // Ensure all column widths are consistent across the worksheet
            double referenceWidth = sheet.Cells.GetColumnWidth(0);
            int totalColumns = sheet.Cells.MaxColumn + 1; // MaxColumn is zero‑based

            for (int col = 0; col < totalColumns; col++)
            {
                sheet.Cells.SetColumnWidth(col, referenceWidth);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
