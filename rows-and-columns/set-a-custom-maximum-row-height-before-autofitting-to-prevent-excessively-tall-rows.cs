// Title: Set a maximum row height after auto‑fitting rows using Aspose.Cells for .NET (C#)
// AI Prompts: Auto‑fit all rows in a worksheet then limit any row height to 30 points with Aspose.Cells in C#. | Apply text wrapping to a cell, auto‑fit rows, and enforce a custom maximum row height before saving the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to cap row height after autofit rows | prevent Excel rows from expanding too much when using AutoFitRows in Aspose.Cells | set a maximum row height limit in a generated workbook with Aspose.Cells for .NET | C# Aspose.Cells set row height ceiling after auto‑fitting rows
// Tags: post‑autofit row height cap Aspose.Cells | configure row height ceiling .NET Excel | cell wrap with row height restriction Aspose.Cells | excel row height cap C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a workbook, adds short and long wrapped text, auto‑fits all rows, then caps any row height exceeding 30 points before saving to output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook(); // For loading: new Workbook("input.xlsx")

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data that may cause tall rows
                sheet.Cells["A1"].PutValue("Short text");
                sheet.Cells["A2"].PutValue("This is a very long text that will wrap and potentially make the row height very large when auto‑fitted.");

                // Enable text wrapping for the long text cell
                Style style = sheet.Cells["A2"].GetStyle();
                style.IsTextWrapped = true;
                sheet.Cells["A2"].SetStyle(style);

                // Define a custom maximum row height (in points)
                double maxRowHeight = 30.0; // Adjust as needed

                // Auto‑fit all rows in the worksheet
                sheet.AutoFitRows();

                // Enforce the maximum row height
                foreach (Row row in sheet.Cells.Rows)
                {
                    if (row.Height > maxRowHeight)
                    {
                        row.Height = maxRowHeight;
                    }
                }

                // Prepare output path and ensure directory exists
                string outputPath = "output.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
