// Title: Validate timeline date ranges in an Excel workbook against a predefined calendar and log inconsistencies with Aspose.Cells (C#)
// Common Searches: C# Aspose.Cells read Excel rows and verify start/end dates are within allowed year ranges | How to log rows with invalid or out-of-range timeline dates from an Excel sheet using Aspose.Cells | Validate that Excel start date is not after end date and fits a predefined calendar with Aspose.Cells in .NET
// Tags: Aspose.Cells validate timeline dates | Excel date consistency check C# | C# write inconsistency log to file | calendar range verification Aspose.Cells | parse Excel date columns C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace TimelineValidator
{
    // The example loads an Excel workbook with Aspose.Cells, iterates each row to parse start and end dates, ensures the start date precedes the end date, checks that the interval falls within predefined calendar ranges (e.g., 2023‑2024), and writes any format errors, reversed dates, or out‑of‑range entries to a text log, optionally including a description column.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file containing the timeline data
            string excelPath = "TimelineData.xlsx";

            // Path to the log file where inconsistencies will be recorded
            string logPath = "Inconsistencies.txt";

            // Ensure the directory for the log file exists
            string logDir = Path.GetDirectoryName(logPath);
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            // Ensure the log file is empty before starting
            File.WriteAllText(logPath, string.Empty);

            // Verify that the Excel file exists before attempting to load it
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Assume the timeline data is in the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range to iterate over rows
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow; // zero‑based index of the last row with data
                int startRow = 1; // assuming row 0 contains headers

                // Column indices (zero‑based). Adjust if your layout differs.
                int startDateCol = 0; // Column A
                int endDateCol = 1;   // Column B
                int descriptionCol = 2; // Column C (optional, for logging)

                // Define the predefined calendar as a list of allowed date ranges
                List<(DateTime Start, DateTime End)> allowedRanges = new List<(DateTime, DateTime)>
                {
                    (new DateTime(2023, 1, 1), new DateTime(2023, 12, 31)),
                    (new DateTime(2024, 1, 1), new DateTime(2024, 12, 31))
                };

                for (int row = startRow; row <= maxRow; row++)
                {
                    // Read start and end dates from the cells
                    object startObj = cells[row, startDateCol].Value;
                    object endObj = cells[row, endDateCol].Value;

                    // Skip rows with missing dates
                    if (startObj == null || endObj == null)
                        continue;

                    // Try to parse the dates
                    if (!DateTime.TryParse(startObj.ToString(), out DateTime startDate) ||
                        !DateTime.TryParse(endObj.ToString(), out DateTime endDate))
                    {
                        LogInconsistency(logPath, row, "Invalid date format.", cells, descriptionCol);
                        continue;
                    }

                    // Validate that start date is not after end date
                    if (startDate > endDate)
                    {
                        LogInconsistency(logPath, row, "Start date is after end date.", cells, descriptionCol);
                        continue;
                    }

                    // Check if the timeline range fits within any allowed calendar range
                    bool isWithinAllowed = false;
                    foreach (var range in allowedRanges)
                    {
                        if (startDate >= range.Start && endDate <= range.End)
                        {
                            isWithinAllowed = true;
                            break;
                        }
                    }

                    if (!isWithinAllowed)
                    {
                        LogInconsistency(logPath, row, "Date range falls outside the predefined calendar.", cells, descriptionCol);
                    }
                }

                Console.WriteLine("Validation complete. Check the log file for any inconsistencies.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }

        /// <param name="logPath">Path to the log file.</param>
        /// <param name="rowIndex">Zero‑based row index where the issue was found.</param>
        /// <param name="message">Description of the inconsistency.</param>
        /// <param name="cells">Reference to the worksheet cells (used to fetch optional description).</param>
        /// <param name="descriptionCol">Column index containing a description or identifier for the timeline entry.</param>
        private static void LogInconsistency(string logPath, int rowIndex, string message, Cells cells, int descriptionCol)
        {
            try
            {
                string description = string.Empty;
                if (descriptionCol >= 0 && cells[rowIndex, descriptionCol].Value != null)
                {
                    description = cells[rowIndex, descriptionCol].Value.ToString();
                }

                string logEntry = $"Row {rowIndex + 1}: {message}";
                if (!string.IsNullOrEmpty(description))
                {
                    logEntry += $" (Description: {description})";
                }

                File.AppendAllText(logPath, logEntry + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write log entry for row {rowIndex + 1}: {ex.Message}");
            }
        }
    }
}
