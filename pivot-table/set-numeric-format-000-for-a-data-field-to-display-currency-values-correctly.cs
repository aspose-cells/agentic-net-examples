// Title: Set column B to currency format "#,##0.00" in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to apply "#,##0.00" numeric format to a column | format Excel cells as currency with thousands separator using Aspose.Cells .NET | apply custom number format to specific cells in a workbook with Aspose.Cells C# | set cell style Custom property for currency display in Aspose.Cells | save modified Excel file after applying numeric format with Aspose.Cells in C#
// Tags: Aspose.Cells set custom numeric format C# | apply currency number format Excel column .NET | cell style Custom property Aspose.Cells | format column B as currency Aspose.Cells | save workbook after style changes Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file with Aspose.Cells, iterates over the first ten rows of column B, assigns the custom numeric format "#,##0.00" to each cell's style to show currency with thousands separators and two decimal places, and then saves the workbook to a new file while handling potential Aspose.Cells and generic exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook. If the file is password‑protected, a CellsException will be thrown.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Apply the numeric format '#,##0.00' to column B (index 1), rows 1‑10.
            for (int row = 0; row < 10; row++)
            {
                Cell cell = sheet.Cells[row, 1]; // Column B
                Style style = cell.GetStyle();
                style.Custom = "#,##0.00";      // Numeric format with thousands separator and two decimals
                cell.SetStyle(style);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles errors such as missing password, corrupted file, etc.
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors.
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
