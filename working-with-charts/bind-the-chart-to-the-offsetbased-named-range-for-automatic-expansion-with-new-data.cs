// Title: Create a dynamic column chart in Aspose.Cells C# by binding it to OFFSET‑based named ranges
// Common Searches: Aspose.Cells C# bind chart series to OFFSET named range for expanding data | How to use the OFFSET function with named ranges in an Aspose.Cells chart | C# Aspose.Cells dynamic chart data range using OFFSET formula | Create an Excel column chart with auto‑growing data range in Aspose.Cells | Aspose.Cells chart XValues from a named range defined by OFFSET in C#
// Tags: OFFSET named range chart binding Aspose.Cells | dynamic chart data range C# Aspose.Cells | auto expanding chart series Aspose.Cells | Excel OFFSET function with Aspose.Cells | column chart from named range C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, defines two OFFSET‑based named ranges (ValuesRange and CategoryRange) that automatically expand as rows are added, adds a column chart, and binds the series values and X‑axis categories to those named ranges before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (Category in column A, Value in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B4"].PutValue(30);

                // ----- Define OFFSET‑based named ranges for dynamic data expansion -----

                // Named range for the values column (B) starting at B2, height determined by COUNTA
                int valuesRangeIdx = workbook.Worksheets.Names.Add("ValuesRange");
                workbook.Worksheets.Names[valuesRangeIdx].RefersTo =
                    "=OFFSET(Sheet1!$B$2,0,0,COUNTA(Sheet1!$B:$B)-1,1)";

                // Named range for the categories column (A) starting at A2, height determined by COUNTA
                int categoriesRangeIdx = workbook.Worksheets.Names.Add("CategoryRange");
                workbook.Worksheets.Names[categoriesRangeIdx].RefersTo =
                    "=OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)";

                // ----- Add a chart and bind it to the OFFSET‑based named ranges -----

                // Add a column chart positioned at row 6, column 1 (zero‑based indices)
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Add the series using the named range for values
                chart.NSeries.Add("{=ValuesRange}", true);

                // Set the category (X‑axis) data using the named range for categories
                chart.NSeries[0].XValues = "{=CategoryRange}";

                // Optional: give the series a name
                chart.NSeries[0].Name = "Sales";

                // Define output file path
                string outputPath = "ChartWithOffsetNamedRange.xlsx";

                // Ensure the directory exists before saving
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
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
