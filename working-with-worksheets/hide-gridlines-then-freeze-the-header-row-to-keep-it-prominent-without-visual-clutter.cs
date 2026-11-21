// Title: Hide worksheet gridlines and freeze the top header row in an Excel file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide worksheet gridlines example | How to freeze the first row in an Excel workbook using Aspose.Cells .NET | Aspose.Cells hide gridlines and keep header row visible | Freeze panes for header row with Aspose.Cells in C# | Remove Excel gridlines programmatically with Aspose.Cells
// Tags: hide worksheet gridlines Aspose.Cells | freeze header row Aspose.Cells | Aspose.Cells gridlines visibility C# | Aspose.Cells freeze panes C# | Excel workbook formatting Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing .xlsx file, disables gridline visibility on the first worksheet, freezes the first row so the header stays in view, and saves the result to a new workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Hide gridlines
                worksheet.IsGridlinesVisible = false;

                // Freeze the header row (first row)
                // Freeze first row, no columns
                worksheet.FreezePanes(1, 0, 1, 0);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
