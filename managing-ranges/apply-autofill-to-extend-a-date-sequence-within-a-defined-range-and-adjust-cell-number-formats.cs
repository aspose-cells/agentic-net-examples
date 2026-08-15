// Title: C# Aspose.Cells AutoFill Date Series and Preserve Number Format
// Description: Demonstrates how to create a workbook, write a short date series (A1:A3), apply the built‑in Excel date format (mm‑dd‑yy, number 14), auto‑fill the series to a larger range (A4:A10) using AutoFillType.Series, and reapply the same number format to the filled cells before saving the file.
// Keywords: Aspose.Cells C# AutoFill date series | preserve date number format Aspose.Cells | extend Excel date range .NET | AutoFillType.Series example | apply style number format Aspose.Cells | C# Excel date autofill | Aspose.Cells date formatting | Workbook date series extension
// Common Searches: Aspose.Cells auto fill date series C# | keep date format after autofill Aspose.Cells | apply number format 14 to cells Aspose.Cells | extend dates A1:A3 to A4:A10 Aspose.Cells | C# example AutoFillType.Series Aspose.Cells
// Developer Intent: Generate a workbook, auto‑fill a date sequence to a defined range, and ensure the same date number format is applied to both source and target cells.
// Use Cases: Create a calendar column by extending an initial three‑day sample while maintaining consistent mm‑dd‑yy formatting. | Build scheduling sheets where dates must continue beyond a starter range and display uniformly. | Populate reporting tables with continuous dates and enforce a single date style across all filled cells.
// AI Prompts: Provide C# code using Aspose.Cells to auto‑fill a date series from A1:A3 to A4:A10 and retain the mm‑dd‑yy format. | Show how to create a style with number format 14 in Aspose.Cells and apply it to source and target ranges after AutoFill. | Explain how to change the date format to a custom pattern after extending a date series with AutoFill in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDateAutoFillDemo
{
    // Demonstrates how to create a workbook, write a short date series (A1:A3), apply the built‑in Excel date format (mm‑dd‑yy, number 14), auto‑fill the series to a larger range (A4:A10) using AutoFillType.Series, and reapply the same number format to the filled cells before saving the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // 1. Populate a short date series as the source range (A1:A3)
                // ------------------------------------------------------------
                cells["A1"].PutValue(new DateTime(2023, 1, 1));
                cells["A2"].PutValue(new DateTime(2023, 1, 2));
                cells["A3"].PutValue(new DateTime(2023, 1, 3));

                // Apply a date number format to the source cells (optional but ensures consistency)
                Style dateStyle = workbook.CreateStyle();
                // Excel date format code 14 corresponds to "mm-dd-yy"
                dateStyle.Number = 14;

                // Use a StyleFlag that applies all style attributes
                StyleFlag flag = new StyleFlag { All = true };

                // Apply the style to the source range
                Aspose.Cells.Range sourceRange = cells.CreateRange("A1:A3");
                sourceRange.ApplyStyle(dateStyle, flag);

                // ------------------------------------------------------------
                // 2. Define source and target ranges
                // ------------------------------------------------------------
                // Source range: A1:A3 (the three dates we just entered)
                // Target range: A4:A10 (where we want the series to continue)
                Aspose.Cells.Range targetRange = cells.CreateRange("A4:A10");

                // ------------------------------------------------------------
                // 3. AutoFill the target range using the Series fill type
                // ------------------------------------------------------------
                sourceRange.AutoFill(targetRange, AutoFillType.Series);

                // ------------------------------------------------------------
                // 4. Ensure the target cells use the same date number format
                // ------------------------------------------------------------
                targetRange.ApplyStyle(dateStyle, flag);

                // ------------------------------------------------------------
                // 5. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "DateAutoFillResult.xlsx";
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
