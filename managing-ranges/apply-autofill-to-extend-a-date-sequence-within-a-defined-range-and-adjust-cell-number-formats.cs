// Title: Extend a date series from A1 to A20 and apply a short date number format with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads (or creates) an Excel workbook, fills cells A1:A20 with consecutive dates starting from the first cell, and formats the entire range with the built‑in short date style using Aspose.Cells. | Show how to define a Style object with Number format 14 in Aspose.Cells, apply it to a specific range, and save the workbook after extending the date column programmatically. | Adapt the example to accept a custom start date and row count, generate the date series accordingly, apply the same short date formatting, and export the file.
// Common Searches: aspocells c# autofill date series to specific row range | apply built‑in short date format to a column using Aspose.Cells .NET | extend existing Excel date column programmatically with Aspose.Cells | create and apply style to a range of cells in Aspose.Cells C# example | save workbook after formatting dates with Aspose.Cells for .NET
// Tags: Aspose.Cells extend date series C# | Aspose.Cells apply short date format range | Aspose.Cells create style number format 14 | Aspose.Cells autofill dates programmatically | Aspose.Cells save workbook after formatting

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDateAutofill
{
    // The sample loads an existing workbook or creates a new one, populates cells A1:A5 with dates, then extends the series to A20 by adding days to the initial date. It creates a Style with the built‑in short date number format (14) and applies this style to the range A1:A20. Finally, it ensures the output directory exists and saves the workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one with sample dates
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    // Populate A1:A5 with a simple date series for demonstration
                    for (int i = 0; i < 5; i++)
                    {
                        ws.Cells[i, 0].PutValue(DateTime.Today.AddDays(i));
                    }
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Extend the date series from A1:A5 to A1:A20 manually
                DateTime startDate = sheet.Cells["A1"].DateTimeValue; // retrieve the first date
                for (int row = 5; row < 20; row++) // rows are zero‑based; row 5 corresponds to A6
                {
                    sheet.Cells[row, 0].PutValue(startDate.AddDays(row));
                }

                // Create a date style (built‑in short date format)
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // Short Date format

                // Apply the style to the extended range (A1:A20)
                Aspose.Cells.Range extendedRange = sheet.Cells.CreateRange(
                    0, // start row
                    0, // start column
                    20, // number of rows (A1 to A20)
                    1   // number of columns
                );

                StyleFlag flag = new StyleFlag { All = true };
                extendedRange.ApplyStyle(dateStyle, flag);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
