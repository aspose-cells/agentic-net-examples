// Title: Auto-fit all columns in an Aspose.Cells worksheet before freezing panes to preserve column widths (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to auto‑fit every column in a worksheet and then freeze the first row, ensuring the column widths stay unchanged. | Generate a snippet that loads an existing Excel file with Aspose.Cells, calls AutoFitColumns on the worksheet, applies FreezePanes, and saves the modified workbook. | Show how to adapt the example to auto‑fit a specific column range before applying FreezePanes in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# auto fit columns then freeze top row without losing width | preserve column width when using FreezePanes in Aspose.Cells .NET | how to call AutoFitColumns before FreezePanes in Aspose.Cells | C# Aspose.Cells example for auto‑fitting all columns and freezing panes
// Tags: AutoFitColumns with FreezePanes Aspose.Cells | preserve column width Aspose.Cells .NET | freeze top row after column auto‑fit C# | Aspose.Cells worksheet column sizing | Excel workbook column auto‑fit before freeze panes

using System;
using System.IO;
using Aspose.Cells;

// The program loads or creates a workbook, accesses the first worksheet, auto‑fits all its columns, freezes the top row using FreezePanes, and saves the modified file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if the file exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Auto‑fit all columns in the worksheet.
            sheet.AutoFitColumns();

            // Freeze panes (example: freeze the top row).
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
