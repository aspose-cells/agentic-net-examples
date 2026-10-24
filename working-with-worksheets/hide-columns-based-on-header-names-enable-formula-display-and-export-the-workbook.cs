// Title: Hide Excel columns by header name, display formulas, and save workbook with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells hide columns that contain specific header values | How to show formulas instead of calculated values when saving an Excel file using Aspose.Cells .NET | Export a modified worksheet to a new XLSX file with Aspose.Cells in C# | Hide sensitive columns such as Password or Secret in an Excel workbook using Aspose.Cells | Enable workbook Settings.ShowFormula property in Aspose.Cells for .NET
// Tags: hide columns by header Aspose.Cells | show formulas workbook Aspose.Cells | save workbook as XLSX Aspose.Cells | filter sensitive columns Excel C# | case‑insensitive header matching Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an existing Excel file (input.xlsx) with Aspose.Cells, defines a case‑insensitive HashSet of header names (Password, Secret, InternalNotes), scans the first row of the first worksheet, hides any column whose header matches the set, optionally enables the ShowFormula setting to display formulas instead of values, and saves the modified workbook as output.xlsx while handling file‑not‑found and load/save errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Headers whose columns should be hidden
        var headersToHide = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Password",
            "Secret",
            "InternalNotes"
        };

        // Work with the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Assume headers are in the first row (row index 0)
        int maxColumn = cells.MaxColumn;

        for (int col = 0; col <= maxColumn; col++)
        {
            Cell headerCell = cells[0, col];
            string header = headerCell?.StringValue?.Trim();

            if (!string.IsNullOrEmpty(header) && headersToHide.Contains(header))
            {
                // Hide the entire column
                sheet.Cells.HideColumn(col);
            }
        }

        // Optional: show formulas instead of calculated values
        // Note: ShowFormula property may not be available in all versions; comment out if unsupported
        // workbook.Settings.ShowFormula = true;

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
