// Title: How to use Aspose.Cells for .NET to apply conditional formatting that colors negative numbers red in an Excel column
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, inserts sample data, adds a CellValue < 0 conditional formatting rule for range A1:A3, and applies a red font style. | Write a C# snippet using Aspose.Cells to define a style with a red font, attach it to a less‑than‑zero condition on a cell range, and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# conditional formatting font color for values less than zero | How to highlight negative numbers in Excel using Aspose.Cells .NET | C# example adding CellValue less than 0 conditional formatting with Aspose | Saving an Excel file with conditional formatting rules using Aspose.Cells | Apply red font style to a range of cells based on value in Aspose.Cells C#
// Tags: Aspose.Cells conditional formatting negative values | Aspose.Cells set font color red | Aspose.Cells CellValue less than zero rule | Aspose.Cells create style for conditional formatting | Aspose.Cells save workbook with formatting

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The program creates a new workbook, adds sample numbers, defines a conditional formatting rule for cells A1:A3 that changes the font color to red when the cell value is less than zero, and saves the file as ConditionalFormatting.xlsx.
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

                // Sample data (can be removed if workbook already contains data)
                sheet.Cells["A1"].PutValue(12);
                sheet.Cells["A2"].PutValue(-7);
                sheet.Cells["A3"].PutValue(3);

                // Add conditional formatting to range A1:A3
                int cfIndex = sheet.ConditionalFormattings.Add();
                var cf = sheet.ConditionalFormattings[cfIndex];

                // Define the area (A1:A3) for the conditional formatting
                CellArea area = new CellArea
                {
                    StartRow = 0,   // Row 1 (zero‑based)
                    StartColumn = 0, // Column A (zero‑based)
                    EndRow = 2,     // Row 3
                    EndColumn = 0   // Column A
                };
                cf.AddArea(area);

                // Condition: cell value less than 0 (negative numbers)
                int conditionIdx = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.LessThan,
                    "0",    // formula1
                    null    // formula2 (not required for this operator)
                );

                // Create a style for the condition (red font)
                Style negativeStyle = workbook.CreateStyle();
                negativeStyle.Font.Color = Color.Red;

                // Apply the style to the condition
                cf[conditionIdx].Style = negativeStyle;

                // Define output file path
                string outputPath = "ConditionalFormatting.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
