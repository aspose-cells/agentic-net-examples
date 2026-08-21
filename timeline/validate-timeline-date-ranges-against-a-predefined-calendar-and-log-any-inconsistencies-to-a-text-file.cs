// Title: C# – Validate Timeline Dates with Aspose.Cells and Log Invalid Entries
// Description: Creates an Excel workbook with a PivotTable and Timeline, checks each ship date against a predefined calendar, writes out‑of‑range dates to a text log, and saves the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells timeline validation | C# date validation Excel | log invalid dates Aspose.Cells | predefined calendar check | PivotTable timeline C# | Excel date consistency
// Common Searches: Aspose.Cells validate timeline dates C# | log dates not in calendar Excel .NET | how to check pivot table dates with Aspose.Cells | timeline control date validation example
// Developer Intent: Ensure that dates displayed in an Aspose.Cells Timeline match a business‑defined calendar and capture any mismatches for review.
// Use Cases: Identify sales records with ship dates outside the approved schedule and generate a log for auditors. | Prevent users from selecting invalid dates in a timeline by cross‑checking with a corporate calendar before workbook distribution. | Automate quality checks on imported Excel data by flagging dates that are not part of a predefined set.
// AI Prompts: Generate code that reads valid dates from a JSON file, validates the Timeline dates, and appends mismatches to a log. | Show how to apply conditional formatting to rows with invalid dates after the validation loop. | Provide a method to export the validation log to CSV instead of plain text.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineDateRangeValidation
{
    // Creates an Excel workbook with a PivotTable and Timeline, checks each ship date against a predefined calendar, writes out‑of‑range dates to a text log, and saves the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Predefined calendar – dates that are considered valid
                HashSet<DateTime> validCalendar = new HashSet<DateTime>
                {
                    new DateTime(2023, 1, 1),
                    new DateTime(2023, 1, 15),
                    new DateTime(2023, 2, 1),
                    new DateTime(2023, 2, 15)
                };

                // -----------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data (date + value)
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header
                sheet.Cells["A1"].PutValue("Ship Date");
                sheet.Cells["B1"].PutValue("Sales");

                // Sample dates – some are valid, some are not
                sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));   // valid
                sheet.Cells["A3"].PutValue(new DateTime(2023, 1, 10));  // invalid
                sheet.Cells["A4"].PutValue(new DateTime(2023, 2, 1));   // valid
                sheet.Cells["A5"].PutValue(new DateTime(2023, 3, 5));   // invalid

                // Sample sales values
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["B4"].PutValue(2000);
                sheet.Cells["B5"].PutValue(2500);

                // -----------------------------------------------------------------
                // 2. Create a PivotTable based on the data (date field will be used for the Timeline)
                // -----------------------------------------------------------------
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Ship Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // The row field contains DateTime values; Aspose.Cells automatically treats it as a date field.
                // No explicit IsDate property is required in recent versions.

                pivot.RefreshData();
                pivot.CalculateData();

                // -----------------------------------------------------------------
                // 3. Add a Timeline control linked to the PivotTable's date field
                // -----------------------------------------------------------------
                int timelineIdx = sheet.Timelines.Add(pivot, "F1", "Ship Date");
                Timeline timeline = sheet.Timelines[timelineIdx];
                timeline.Caption = "Ship Date Timeline";

                // -----------------------------------------------------------------
                // 4. Validate each date in the source column against the predefined calendar
                //    Log any inconsistencies to a text file.
                // -----------------------------------------------------------------
                string logPath = "TimelineDateValidationLog.txt";
                using (StreamWriter writer = new StreamWriter(logPath, false))
                {
                    int dateColumnIndex = 0;   // Column A (0‑based)
                    int firstDataRow = 1;      // Row after header (0‑based)

                    for (int row = firstDataRow; row <= sheet.Cells.MaxDataRow; row++)
                    {
                        object cellValue = sheet.Cells[row, dateColumnIndex].Value;
                        if (cellValue is DateTime date)
                        {
                            if (!validCalendar.Contains(date.Date))
                            {
                                writer.WriteLine($"Row {row + 1}: Date {date:yyyy-MM-dd} is not in the predefined calendar.");
                            }
                        }
                    }
                }

                // -----------------------------------------------------------------
                // 5. Save the workbook (the Timeline is now part of the sheet)
                // -----------------------------------------------------------------
                string outputPath = "TimelineDateValidationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'. Log written to '{logPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
