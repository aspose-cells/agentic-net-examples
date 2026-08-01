// Title: Aspose.Cells for .NET: Export to Tab‑Delimited Text with Fixed‑Point Numbers Using TxtSaveOptions
// Description: C# example that creates a workbook, applies a two‑decimal style to cells A1:A3, sets TxtSaveOptions to use the DisplayStyle strategy, and saves the data as a tab‑separated text file. The resulting file contains numbers in plain decimal form instead of scientific notation.
// Keywords: Aspose.Cells | .NET | TxtSaveOptions | CellValueFormatStrategy.DisplayStyle | non exponential | fixed point format | tab delimited export | C# example | prevent scientific notation | Excel to txt
// Common Searches: Aspose.Cells export to txt without scientific notation | TxtSaveOptions DisplayStyle C# | force fixed point format when saving Excel as text | prevent exponential numbers in Aspose.Cells output | tab separated export with number formatting Aspose.Cells
// Developer Intent: Export numeric cells to a text file while preserving the on‑screen fixed‑point formatting and avoiding scientific notation.
// Use Cases: Generate tab‑delimited data files for downstream systems that require plain decimal numbers. | Create reports where numeric precision must match the workbook’s displayed format. | Automate data extraction from Excel worksheets without losing formatting for large or small values.
// AI Prompts: Show how to apply a custom number format to a range and save the workbook as tab‑delimited text with non‑exponential numbers using Aspose.Cells. | Provide a C# snippet that uses TxtSaveOptions and CellValueFormatStrategy.DisplayStyle to export fixed‑point numbers. | Explain how to configure SaveOptions in Aspose.Cells to enforce display formatting for all numeric cells when saving to a text file.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// C# example that creates a workbook, applies a two‑decimal style to cells A1:A3, sets TxtSaveOptions to use the DisplayStyle strategy, and saves the data as a tab‑separated text file. The resulting file contains numbers in plain decimal form instead of scientific notation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with numeric values that would normally be shown in exponential form
            sheet.Cells["A1"].PutValue(123456789);          // Large integer
            sheet.Cells["A2"].PutValue(0.00000123);        // Small decimal
            sheet.Cells["A3"].PutValue(1.23e+10);          // Scientific notation

            // Apply a number format that forces a fixed‑point display (e.g., two decimal places)
            Style numberStyle = workbook.CreateStyle();
            numberStyle.Number = 2; // "0.00" format

            // Set the style for the range A1:A3
            AsposeRange range = sheet.Cells.CreateRange("A1:A3");
            range.SetStyle(numberStyle);

            // Configure TxtSaveOptions to use the display style (non‑exponential) for all numeric cells
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = '\t',                                 // Tab‑separated values
                FormatStrategy = CellValueFormatStrategy.DisplayStyle // Use formatted (display) values
            };

            // Save the workbook using the configured options
            workbook.Save("NonExponentialOutput.txt", saveOptions);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
