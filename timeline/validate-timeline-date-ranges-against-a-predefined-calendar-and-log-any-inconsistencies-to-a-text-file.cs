// Title: Validate timeline start and end dates against a calendar worksheet and log errors using Aspose.Cells in C#
// AI Prompts: Create a C# console program that opens an Excel workbook containing 'Timeline' and 'Calendar' sheets, builds a HashSet of dates from column A of the Calendar sheet, checks each row in the Timeline sheet for valid DateTime values, ensures the start date is not later than the end date, and writes any violations to a specified text log. | Modify an existing Aspose.Cells script to handle missing worksheets, verify date formats, compare dates against a predefined calendar range, and output detailed discrepancy messages (including row number and ID) to a log file.
// Common Searches: aspnet c# how to check if Excel dates exist in a separate calendar sheet using Aspose.Cells | c# Aspose.Cells validate start date before end date in timeline worksheet | write validation log for Excel date range errors with Aspose.Cells | load Excel workbook and verify date consistency across two sheets in C# | fast lookup of calendar dates in C# using Aspose.Cells
// Tags: Aspose.Cells read Excel dates into HashSet | compare timeline entries to master calendar using Aspose.Cells | log Excel validation errors to text file | ensure chronological order of timeline dates with Aspose.Cells | detect missing calendar dates in Excel timeline

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace TimelineValidation
{
    // A C# console application that loads 'TimelineData.xlsx', extracts valid dates from the 'Calendar' sheet into a HashSet, iterates through rows of the 'Timeline' sheet, validates date formats, checks that start dates precede end dates, confirms each date exists in the calendar, and records any inconsistencies to 'TimelineValidationLog.txt'.
    class TimelineValidator
    {
        static void Main()
        {
            // Paths to the Excel file and the log file
            string excelPath = @"C:\Data\TimelineData.xlsx";
            string logPath = @"C:\Data\TimelineValidationLog.txt";

            // Ensure the Excel file exists before loading
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Retrieve required worksheets
                Worksheet timelineSheet = workbook.Worksheets["Timeline"];
                Worksheet calendarSheet = workbook.Worksheets["Calendar"];

                if (timelineSheet == null || calendarSheet == null)
                {
                    Console.WriteLine("Error: Required worksheets 'Timeline' or 'Calendar' not found.");
                    return;
                }

                // Build a set of valid dates from the calendar sheet (dates in column A)
                HashSet<DateTime> validDates = new HashSet<DateTime>();
                int calendarRowCount = calendarSheet.Cells.MaxDataRow + 1;
                for (int i = 0; i < calendarRowCount; i++)
                {
                    object cellValue = calendarSheet.Cells[i, 0].Value;
                    if (cellValue is DateTime dt)
                    {
                        validDates.Add(dt.Date); // store date only
                    }
                }

                // Ensure the log directory exists
                string logDir = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                // Prepare the log file (overwrite any existing content)
                using (StreamWriter logWriter = new StreamWriter(logPath, false))
                {
                    // Iterate through timeline rows (header in row 0, data starts at row 1)
                    int timelineRowCount = timelineSheet.Cells.MaxDataRow + 1;
                    for (int row = 1; row < timelineRowCount; row++)
                    {
                        // Read ID (optional, for reporting)
                        string id = timelineSheet.Cells[row, 0].StringValue;

                        // Read start and end dates (columns B and C)
                        object startObj = timelineSheet.Cells[row, 1].Value;
                        object endObj = timelineSheet.Cells[row, 2].Value;

                        // Validate that both cells contain dates
                        if (!(startObj is DateTime startDate) || !(endObj is DateTime endDate))
                        {
                            logWriter.WriteLine($"Row {row + 1} (ID: {id}) - Invalid date format.");
                            continue;
                        }

                        // Normalize to date only
                        startDate = startDate.Date;
                        endDate = endDate.Date;

                        // Check logical order
                        if (startDate > endDate)
                        {
                            logWriter.WriteLine($"Row {row + 1} (ID: {id}) - Start date {startDate:d} is after end date {endDate:d}.");
                        }

                        // Verify each date falls within the predefined calendar
                        if (!validDates.Contains(startDate))
                        {
                            logWriter.WriteLine($"Row {row + 1} (ID: {id}) - Start date {startDate:d} is not in the calendar.");
                        }

                        if (!validDates.Contains(endDate))
                        {
                            logWriter.WriteLine($"Row {row + 1} (ID: {id}) - End date {endDate:d} is not in the calendar.");
                        }
                    }
                }

                Console.WriteLine($"Validation completed. Log written to '{logPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
