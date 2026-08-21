// Title: Apply template date format to generated dates with Aspose.Cells CopyStyle in C#
// Description: Shows how to load or create a template workbook that defines a short date style, copy that style to a target range (B2:B6) in a new workbook using Aspose.Cells CopyStyle, fill the range with sequential DateTime values, and save the result so the dates retain the template's locale‑specific formatting.
// Keywords: Aspose.Cells CopyStyle | C# date format | locale specific date formatting | copy cell style between workbooks | template workbook date style | inherit date formatting Aspose.Cells | Excel date style .NET | range style copy example | Aspose.Cells date formatting
// Common Searches: Aspose.Cells CopyStyle date format C# | how to copy date style from one workbook to another Aspose.Cells | inherit locale specific date format in generated Excel file | copy cell style before populating values Aspose.Cells | apply template date format to new workbook using Aspose.Cells
// Developer Intent: Copy a date style from a template range and apply it to a generated date range so the output uses the same locale‑specific format.
// Use Cases: Generate a series of dates in a report workbook that match the date format defined in a reusable template. | Create a temporary template with a custom date style and reuse it across multiple Excel exports. | Ensure consistent date formatting when populating large data sets by copying the style before inserting values.
// AI Prompts: Provide a C# example that copies a custom date format from a template workbook to a target range using Aspose.Cells CopyStyle. | Show how to preserve locale‑specific date formatting when generating dates in a new Excel file with Aspose.Cells. | Explain the steps to create a template workbook with a short date style, copy the style to another workbook, and fill the range with sequential dates in C#.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCopyStyleDateDemo
{
    // Shows how to load or create a template workbook that defines a short date style, copy that style to a target range (B2:B6) in a new workbook using Aspose.Cells CopyStyle, fill the range with sequential DateTime values, and save the result so the dates retain the template's locale‑specific formatting.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook that contains the desired date format.
                const string templatePath = "TemplateWithDateFormat.xlsx";

                // Load the template workbook if it exists; otherwise create a temporary one with a date style.
                Workbook templateWorkbook;
                if (File.Exists(templatePath))
                {
                    templateWorkbook = new Workbook(templatePath);
                }
                else
                {
                    // Create a workbook and apply a built‑in date format to cell A1.
                    templateWorkbook = new Workbook();
                    Worksheet tempSheet = templateWorkbook.Worksheets[0];
                    Cell tempCell = tempSheet.Cells["A1"];
                    tempCell.PutValue(DateTime.Now);
                    Style dateStyle = tempCell.GetStyle();
                    dateStyle.Number = 14; // Built‑in short date format.
                    tempCell.SetStyle(dateStyle);
                }

                Worksheet templateSheet = templateWorkbook.Worksheets[0];
                // Create a range that refers to cell A1 in the template.
                Aspose.Cells.Range templateDateRange = templateSheet.Cells.CreateRange("A1");

                // Create a new workbook where dates will be generated.
                Workbook resultWorkbook = new Workbook();
                Worksheet resultSheet = resultWorkbook.Worksheets[0];

                // Define the target range for generated dates (B2:B6).
                Aspose.Cells.Range targetDateRange = resultSheet.Cells.CreateRange(1, 1, 5, 1); // rows 2‑6, column B

                // Copy the date style from the template range to the target range.
                targetDateRange.CopyStyle(templateDateRange);

                // Populate the target range with date values.
                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 5; i++)
                {
                    // Cells are accessed via zero‑based row and column indexes.
                    Cell cell = resultSheet.Cells[1 + i, 1]; // B2, B3, …
                    cell.PutValue(startDate.AddDays(i));
                    // Style already copied; no further action needed.
                }

                // Save the resulting workbook.
                const string resultPath = "GeneratedDatesWithTemplateStyle.xlsx";
                resultWorkbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
