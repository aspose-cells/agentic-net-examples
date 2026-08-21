// Title: C# Aspose.Cells example: Generate Excel report of Japanese‑formatted dates with original Gregorian values
// Description: A complete Aspose.Cells for .NET sample that creates a workbook, populates cells with Gregorian DateTime values, sets the workbook region to Japan, applies a custom Japanese date format (e.g., "yyyy年M月d日"), and builds a separate worksheet listing each date cell’s address, the original Gregorian value, and the Japanese‑formatted string. The result is saved as JapaneseDateReport.xlsx.
// Keywords: Aspose.Cells | C# | Japanese date format | Excel localization | Gregorian to Japanese conversion | region Japan | custom number format | date conversion report | Excel automation | globalization | localization | Japanese era dates | workbook generation
// Common Searches: Aspose.Cells convert dates to Japanese format | C# generate Excel report of Japanese dates | apply Japanese locale in Aspose.Cells | list original and Japanese dates in Excel using Aspose | create date conversion report with Aspose.Cells .NET
// Developer Intent: Produce an Excel workbook that scans all DateTime cells, converts each to the Japanese calendar format, and records the cell address, original Gregorian DateTime, and formatted string on a dedicated report sheet.
// Use Cases: Audit date‑field localization for a Japanese market release while preserving original timestamps for traceability. | Prepare financial statements that require Japanese era date formatting but must retain Gregorian dates for regulatory compliance. | Document a migration of date formats by generating a side‑by‑side report of source and localized values. | Automate generation of localized Excel reports for multinational teams needing both native and universal date representations.
// AI Prompts: Write C# code using Aspose.Cells that iterates through a worksheet, applies the custom Japanese date format "[$-F800]yyyy年M月d日" to every DateTime cell, and logs the cell address, original Gregorian value, and formatted result on a new report sheet. | Explain how to set the workbook region to Japan in Aspose.Cells and why the custom number format "[$-F800]yyyy年M月d日" displays dates in the Japanese calendar style. | Provide a unit test for the Japanese date report that verifies the report sheet contains correct cell addresses, original DateTime objects, and properly formatted Japanese strings. | Suggest best practices for handling time‑zone differences when converting Gregorian dates to Japanese formatted dates in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsJapaneseDateReport
{
    // A complete Aspose.Cells for .NET sample that creates a workbook, populates cells with Gregorian DateTime values, sets the workbook region to Japan, applies a custom Japanese date format (e.g., "yyyy年M月d日"), and builds a separate worksheet listing each date cell’s address, the original Gregorian value, and the Japanese‑formatted string. The result is saved as JapaneseDateReport.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (source data)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Populate sample Gregorian dates in column A
            for (int i = 0; i < 5; i++)
            {
                // PutValue automatically stores the value as a DateTime
                sourceSheet.Cells[i, 0].PutValue(DateTime.Now.AddDays(i));
            }

            // Set the workbook's regional settings to Japan
            workbook.Settings.Region = CountryCode.Japan;

            // Add a new worksheet for the report
            int reportIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportIndex];
            reportSheet.Name = "JapaneseDateReport";

            // Write report headers
            reportSheet.Cells[0, 0].PutValue("Cell Address");
            reportSheet.Cells[0, 1].PutValue("Original Gregorian Value");
            reportSheet.Cells[0, 2].PutValue("Japanese Formatted Value");

            int reportRow = 1; // start after header

            // Determine the used range in the source sheet
            int maxRow = sourceSheet.Cells.MaxDataRow;
            int maxCol = sourceSheet.Cells.MaxDataColumn;

            // Iterate through all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sourceSheet.Cells[row, col];

                    // Process only cells that contain a DateTime value
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        // Preserve the original Gregorian DateTime
                        DateTime originalDate = cell.DateTimeValue;

                        // Apply Japanese date format (e.g., "2023年5月15日")
                        Style style = cell.GetStyle();
                        style.Custom = "[$-F800]yyyy年m月d日";
                        cell.SetStyle(style);

                        // Record the conversion in the report sheet
                        reportSheet.Cells[reportRow, 0].PutValue(cell.Name);                     // e.g., "A1"
                        reportSheet.Cells[reportRow, 1].PutValue(originalDate);                // original Gregorian
                        reportSheet.Cells[reportRow, 2].PutValue(cell.StringValue);           // Japanese formatted string
                        reportRow++;
                    }
                }
            }

            // Save the workbook with the report
            workbook.Save("JapaneseDateReport.xlsx");
        }
    }
}
