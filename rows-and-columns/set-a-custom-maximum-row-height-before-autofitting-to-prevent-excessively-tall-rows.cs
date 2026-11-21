// Title: C# – Set a maximum row height after AutoFitRows to prevent overly tall rows with Aspose.Cells
// Common Searches: how to restrict row height after calling AutoFitRows in Aspose.Cells for .NET | Aspose.Cells C# limit row height when text wrapping creates tall rows | set maximum row height programmatically after auto‑fit rows in Excel using Aspose.Cells | prevent rows from expanding beyond a certain height after AutoFitRows in C# | apply custom row height cap after AutoFitRows with Aspose.Cells .NET
// Tags: auto‑fit rows with maximum height Aspose.Cells | c# limit row height Aspose.Cells | enforce row height cap after AutoFitRows | wrap text and cap row height Aspose.Cells | custom row height constraint Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, writes short and long text, enables text wrapping for the long cell, auto‑fits all rows, then iterates through each row to enforce a maximum height of 30 points before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load existing workbook if present; otherwise create a new one
                Workbook workbook;
                string inputPath = "input.xlsx";

                if (File.Exists(inputPath))
                {
                    try
                    {
                        workbook = new Workbook(inputPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to load '{inputPath}': {ex.Message}");
                        workbook = new Workbook();
                    }
                }
                else
                {
                    workbook = new Workbook();
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Sample data
                sheet.Cells["A1"].PutValue("Short text");
                sheet.Cells["A2"].PutValue("This is a very long text that will normally cause the row to become very tall when auto‑fitted.");

                // Enable text wrapping for the long cell
                Style wrapStyle = sheet.Cells["A2"].GetStyle();
                wrapStyle.IsTextWrapped = true; // Correct property name
                sheet.Cells["A2"].SetStyle(wrapStyle);

                // Maximum allowed row height (points)
                double maxRowHeight = 30.0;

                // Auto‑fit rows based on content
                sheet.AutoFitRows();

                // Enforce maximum height
                int lastRow = sheet.Cells.MaxDataRow;
                for (int i = 0; i <= lastRow; i++)
                {
                    Row row = sheet.Cells.Rows[i];
                    if (row.Height > maxRowHeight)
                    {
                        row.Height = maxRowHeight;
                        // Setting Height automatically marks the row as having a custom height
                    }
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
