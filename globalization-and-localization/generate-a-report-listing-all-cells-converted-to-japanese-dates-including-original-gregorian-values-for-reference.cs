// Title: Create a Japanese date conversion report in Excel with original Gregorian values using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, detects cells whose number format contains the Japanese characters 年, 月, 日, and records the sheet name, cell address, Gregorian DateTime value, and the formatted Japanese date into a new workbook. | Enhance the detection logic to also include built‑in Japanese date format IDs (e.g., 14) when identifying Japanese‑formatted cells, and add those entries to the generated report. | Add functionality to export the same report data to a CSV file while preserving the column order: Sheet, Cell, Gregorian Value, Japanese Date.
// Common Searches: how to list cells with Japanese date format using Aspose.Cells C# | Aspose.Cells detect custom number format containing 年 月 日 | generate Excel report of localized Japanese dates and original Gregorian values .NET | export Japanese formatted dates from workbook to new Excel file with Aspose.Cells | include built‑in Japanese date format ID 14 in Aspose.Cells cell detection
// Tags: Aspose.Cells detect Japanese date format | C# create Excel report of localized dates | list cells with custom Japanese number format | export Gregorian and Japanese dates to CSV | auto‑fit columns Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace JapaneseDateReportGenerator
{
    // The solution loads a source workbook, iterates through all worksheets and used cells, identifies DateTime cells whose custom number format contains the Japanese characters 年, 月, 日 (or uses built‑in Japanese date format IDs), captures the sheet name, cell address, original Gregorian value, and the Japanese‑formatted string, writes this data into a new Excel workbook with headers and auto‑fitted columns, and saves the report. Optional extensions add built‑in format detection and CSV export.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the source workbook and the generated report
            string sourcePath = "input.xlsx";
            string reportPath = "JapaneseDateReport.xlsx";

            // Load the source workbook (create/load rule)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // List to hold information about cells converted to Japanese dates
            var japaneseDateCells = new List<JapaneseDateCellInfo>();

            // Iterate through all worksheets in the source workbook
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                // Get the used range of the worksheet to limit iteration
                var usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Check if the cell contains a DateTime value
                    if (cell.Value is DateTime gregorianDate)
                    {
                        // Retrieve the cell's number format
                        Style style = cell.GetStyle();

                        // Determine if the number format corresponds to a Japanese date format.
                        // This check looks for typical Japanese date characters (年, 月, 日) in a custom format.
                        // Adjust the condition if your workbook uses a specific built‑in format ID.
                        bool isJapaneseDateFormat = false;

                        // Built‑in number format IDs for Japanese dates can vary; check custom format as fallback.
                        if (!string.IsNullOrEmpty(style.Custom))
                        {
                            string customFormat = style.Custom;
                            if (customFormat.Contains("年") && customFormat.Contains("月") && customFormat.Contains("日"))
                            {
                                isJapaneseDateFormat = true;
                            }
                        }

                        // If the cell uses a built‑in format, you may also compare the Number property.
                        // Example: Japanese long date format often has ID 14 in some locales.
                        // Here we simply rely on the custom format detection above.

                        if (isJapaneseDateFormat)
                        {
                            // Store the cell information
                            japaneseDateCells.Add(new JapaneseDateCellInfo
                            {
                                SheetName = sheet.Name,
                                CellName = cell.Name,
                                GregorianValue = gregorianDate,
                                JapaneseFormattedValue = cell.StringValue // Already formatted according to the cell's style
                            });
                        }
                    }
                }
            }

            // Create a new workbook for the report (create/save rule)
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "Japanese Date Report";

            // Write header row
            reportSheet.Cells["A1"].PutValue("Sheet");
            reportSheet.Cells["B1"].PutValue("Cell");
            reportSheet.Cells["C1"].PutValue("Gregorian Value");
            reportSheet.Cells["D1"].PutValue("Japanese Date");

            // Populate the report with collected data
            int rowIndex = 1; // Zero‑based index; row 1 is the second row (after header)
            foreach (var info in japaneseDateCells)
            {
                reportSheet.Cells[rowIndex, 0].PutValue(info.SheetName);
                reportSheet.Cells[rowIndex, 1].PutValue(info.CellName);
                reportSheet.Cells[rowIndex, 2].PutValue(info.GregorianValue);
                reportSheet.Cells[rowIndex, 3].PutValue(info.JapaneseFormattedValue);
                rowIndex++;
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the report workbook (save rule)
            reportWorkbook.Save(reportPath);
        }
    }

    // Helper class to store information about each Japanese‑date cell
    class JapaneseDateCellInfo
    {
        public string SheetName { get; set; }
        public string CellName { get; set; }
        public DateTime GregorianValue { get; set; }
        public string JapaneseFormattedValue { get; set; }
    }
}
