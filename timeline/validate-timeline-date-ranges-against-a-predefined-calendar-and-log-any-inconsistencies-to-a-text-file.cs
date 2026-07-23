// Title: C# – Validate Aspose.Cells Timeline Dates Against a Calendar and Log Inconsistencies
// Description: Creates a workbook with sample dates, builds a pivot table, adds a timeline linked to the Date field, defines an allowed date range (Jan 1‑31 2023), then checks the timeline's start date and every source‑data date against this range, writing any out‑of‑range entries to a text log before saving the file.
// Keywords: Aspose.Cells timeline validation | C# date range check Excel | log invalid dates Aspose.Cells | timeline start date verification | Excel pivot table date filter | write validation log C# | predefined calendar Aspose.Cells
// Common Searches: Aspose.Cells validate timeline dates C# | How to log dates outside a calendar in Excel using Aspose.Cells | Check timeline start date against custom range .NET | Create validation log for Excel dates with Aspose.Cells | C# example timeline date range enforcement
// Developer Intent: Check that all dates used by an Aspose.Cells timeline and its source data fall within a specified calendar range and record any violations.
// Use Cases: Identify and report dates outside a fiscal period in workbooks that use timeline filters. | Prevent users from selecting out‑of‑scope dates in a project schedule timeline. | Generate an audit file of data‑entry errors before publishing the Excel file. | Automate compliance checks for regulatory reporting periods in Excel dashboards.
// AI Prompts: Generate C# code that opens an existing workbook, adds a timeline to a pivot table, validates each date against a HashSet of allowed dates, and writes mismatches to a text file. | Explain how to retrieve the StartDate and EndDate properties of an Aspose.Cells Timeline and compare them with a custom date range. | Suggest enhancements to handle time‑zone differences and duplicate date entries while logging inconsistencies. | Provide guidance on optimizing the validation loop for large datasets using Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Creates a workbook with sample dates, builds a pivot table, adds a timeline linked to the Date field, defines an allowed date range (Jan 1‑31 2023), then checks the timeline's start date and every source‑data date against this range, writing any out‑of‑range entries to a text log before saving the file.
class TimelineDateValidator
{
    static void Main()
    {
        try
        {
            // -------------------- Create workbook and sample data --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");

            // Sample dates (some inside, some outside the allowed calendar)
            DateTime[] sampleDates = {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 10),
                new DateTime(2023, 2, 1)   // outside allowed range
            };
            int[] sampleValues = { 100, 200, 150, 300 };

            for (int i = 0; i < sampleDates.Length; i++)
            {
                // Date column
                cells[i + 2, 0].PutValue(sampleDates[i]);
                // Ensure the cell is recognized as a date (optional style)
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "yyyy-mm-dd";
                cells[i + 2, 0].SetStyle(dateStyle);

                // Value column
                cells[i + 2, 1].PutValue(sampleValues[i]);
            }

            // -------------------- Create PivotTable --------------------
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:B5", "D1", "Pivot1");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------- Add Timeline linked to the Date field --------------------
            Timeline timeline = null;
            try
            {
                sheet.Timelines.Add(pivot, "F1", "Date");
                timeline = sheet.Timelines[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add timeline: {ex.Message}");
                // Continue without timeline; validation will be limited to source data.
            }

            // -------------------- Define predefined calendar (allowed dates) --------------------
            HashSet<DateTime> allowedDates = new HashSet<DateTime>();
            DateTime calendarStart = new DateTime(2023, 1, 1);
            DateTime calendarEnd = new DateTime(2023, 1, 31);
            for (DateTime d = calendarStart; d <= calendarEnd; d = d.AddDays(1))
            {
                allowedDates.Add(d.Date);
            }

            // -------------------- Validate timeline and source data --------------------
            string logFilePath = "TimelineValidationLog.txt";
            using (StreamWriter log = new StreamWriter(logFilePath, false))
            {
                // Validate Timeline start date (if timeline was created)
                if (timeline != null && !allowedDates.Contains(timeline.StartDate.Date))
                {
                    log.WriteLine($"[Inconsistency] Timeline start date {timeline.StartDate:d} is outside the allowed calendar.");
                }

                // Validate each date in the source data column
                int lastDataRow = cells.MaxDataRow;
                for (int row = 1; row <= lastDataRow; row++) // data starts at row index 1 (A2)
                {
                    object cellValue = cells[row, 0].Value; // Date column (A)
                    if (cellValue is DateTime dt)
                    {
                        if (!allowedDates.Contains(dt.Date))
                        {
                            log.WriteLine($"[Inconsistency] Row {row + 1} contains date {dt:d} outside the allowed calendar.");
                        }
                    }
                }
            }

            // -------------------- Save workbook --------------------
            string outputPath = "TimelineValidated.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }

            Console.WriteLine($"Validation log written to {logFilePath}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Unexpected error: {e.Message}");
        }
    }
}
