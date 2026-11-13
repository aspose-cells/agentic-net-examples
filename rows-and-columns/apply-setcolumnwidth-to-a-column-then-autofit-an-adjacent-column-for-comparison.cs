// Title: Set a fixed column width and auto‑fit an adjacent column in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set column width to a specific number of characters | auto‑fit a column after manually setting width of another column with Aspose.Cells | compare fixed column width versus auto‑fit column in Aspose.Cells example | using SetColumnWidth and AutoFitColumn together in a .NET Excel export | C# code sample for column sizing differences in Aspose.Cells workbook
// Tags: SetColumnWidth Aspose.Cells C# | AutoFitColumn Aspose.Cells C# | manual column width vs auto‑fit Aspose.Cells | Excel column sizing programmatically .NET | column width comparison Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, sets column A (index 0) to a width of 20 characters with SetColumnWidth, auto‑fits column B (index 1) based on its content using AutoFitColumn, and saves the file as 'ColumnWidthComparison.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set column A (index 0) width to 20 characters
                sheet.Cells.SetColumnWidth(0, 20);

                // Auto‑fit column B (index 1) based on its content
                sheet.AutoFitColumn(1);

                // Output file path
                string outputPath = "ColumnWidthComparison.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
