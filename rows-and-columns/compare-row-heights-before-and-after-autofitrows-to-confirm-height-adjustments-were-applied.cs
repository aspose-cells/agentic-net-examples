// Title: How to compare Excel row heights before and after AutoFitRows using Aspose.Cells for .NET
// Common Searches: C# code to capture original row heights and detect changes after Aspose.Cells AutoFitRows | verify if AutoFitRows modifies row height in an Aspose.Cells workbook | compare before and after row height values using Aspose.Cells .NET API | log row height differences after applying AutoFitRows in C# | determine rows that changed height after AutoFitRows with Aspose.Cells
// Tags: Aspose.Cells AutoFitRows row height measurement | original vs adjusted row height Aspose.Cells | C# Excel row height change detection | recording row heights before AutoFitRows | row height validation in Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace RowHeightAutoFitDemoApp
{
    // The example creates a workbook, writes text of varying length, enables text wrapping, records each row's initial Height, runs AutoFitRows to adjust heights, records the new Height values, compares the before‑and‑after heights, outputs any changes, and saves the file.
    class RowHeightAutoFitDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with sample text
                sheet.Cells["A1"].PutValue("Short text");
                sheet.Cells["A2"].PutValue("This is a longer piece of text that should cause the row to increase in height when AutoFitRows is applied.");
                sheet.Cells["A3"].PutValue("Another short text");

                // Enable text wrapping for the second row (index 1)
                Style wrapStyle = workbook.CreateStyle();
                wrapStyle.IsTextWrapped = true;
                StyleFlag flag = new StyleFlag { All = true };
                sheet.Cells.Rows[1].ApplyStyle(wrapStyle, flag);

                // Determine the number of rows to evaluate
                int rowCount = sheet.Cells.Rows.Count;

                // Store original row heights
                double[] originalHeights = new double[rowCount];
                for (int i = 0; i < rowCount; i++)
                {
                    originalHeights[i] = sheet.Cells.Rows[i].Height;
                }

                // AutoFit rows based on content
                sheet.AutoFitRows();

                // Store new row heights after AutoFitRows
                double[] newHeights = new double[rowCount];
                for (int i = 0; i < rowCount; i++)
                {
                    newHeights[i] = sheet.Cells.Rows[i].Height;
                }

                // Output comparison of heights
                Console.WriteLine("Row Height Comparison (Before vs After AutoFitRows):");
                for (int i = 0; i < rowCount; i++)
                {
                    if (Math.Abs(originalHeights[i] - newHeights[i]) > 0.001)
                    {
                        Console.WriteLine($"Row {i + 1}: Height changed from {originalHeights[i]} to {newHeights[i]}");
                    }
                    else
                    {
                        Console.WriteLine($"Row {i + 1}: Height unchanged at {originalHeights[i]}");
                    }
                }

                // Save the workbook
                string outputPath = "RowHeightAutoFitDemo.xlsx";
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
