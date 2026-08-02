// Title: C# – Add a Timeline to a Pivot Table and Apply a Custom Date‑Range Filter with Aspose.Cells
// Description: Demonstrates how to build a workbook, insert sample Date and Sales data, create a PivotTable, add a Timeline control linked to the Date field, set the timeline's visible start date, apply a PivotField date‑between filter (e.g., Jan 1 2023 – Apr 30 2023), and save the result as TimelineFilterDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | PivotTable timeline | Excel timeline control | date range filter | PivotField.FilterByDate | custom start date timeline | date between filter Aspose.Cells | automate Excel reports .NET | timeline filter example
// Common Searches: how to add a timeline to a pivot table using Aspose.Cells C# | set custom start date for timeline control Aspose.Cells | apply date between filter on pivot rows Aspose.Cells .NET | Aspose.Cells timeline filter example | C# code for pivot table date filter with timeline
// Developer Intent: Programmatically attach a Timeline to a PivotTable and restrict the pivot rows to a defined date interval.
// Use Cases: Create a sales dashboard where the timeline shows Jan‑May 2023 while the pivot displays only Jan‑Apr 2023 data. | Build an interactive Excel report that lets users slide a timeline to view sales for any custom period. | Automate monthly reporting by applying a date‑between filter to a pivot field and saving the filtered workbook.
// AI Prompts: Generate C# code that lets end‑users change the timeline's start and end dates at runtime with Aspose.Cells. | Show how to add multiple timelines for different date columns in the same workbook using Aspose.Cells. | Explain how to read the currently selected start and end dates from a Timeline object in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineFilterDemo
{
    // Demonstrates how to build a workbook, insert sample Date and Sales data, create a PivotTable, add a Timeline control linked to the Date field, set the timeline's visible start date, apply a PivotField date‑between filter (e.g., Jan 1 2023 – Apr 30 2023), and save the result as TimelineFilterDemo.xlsx using Aspose.Cells for .NET.
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

                // Populate sample data with a Date column and a Sales column
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Sales");

                cells["A2"].PutValue(new DateTime(2022, 12, 15));
                cells["B2"].PutValue(1200);
                cells["A3"].PutValue(new DateTime(2023, 1, 10));
                cells["B3"].PutValue(1500);
                cells["A4"].PutValue(new DateTime(2023, 2, 5));
                cells["B4"].PutValue(1800);
                cells["A5"].PutValue(new DateTime(2023, 3, 20));
                cells["B5"].PutValue(2100);
                cells["A6"].PutValue(new DateTime(2023, 4, 25));
                cells["B6"].PutValue(2400);
                cells["A7"].PutValue(new DateTime(2023, 5, 30));
                cells["B7"].PutValue(2700);

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the Date field to the Row area and the Sales field to the Data area
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table so that it contains data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline control linked to the Date field of the pivot table
                // The timeline will be placed starting at cell E1
                int timelineIdx = sheet.Timelines.Add(pivot, "E1", "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];

                // Set the visible start date of the timeline to a custom period (e.g., Jan 1, 2023)
                timeline.StartDate = new DateTime(2023, 1, 1);

                // Optionally, set the caption for better UI clarity
                timeline.Caption = "Sales Timeline (Jan‑May 2023)";

                // Apply a date filter on the pivot field to show only dates between Jan 1 and Apr 30, 2023
                PivotField dateField = pivot.RowFields[0];
                dateField.FilterByDate(PivotFilterType.DateBetween,
                                       new DateTime(2023, 1, 1),
                                       new DateTime(2023, 4, 30));

                // Save the workbook with the timeline and filter applied
                workbook.Save("TimelineFilterDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
