// Title: How to add a month-level timeline filter with a custom start date to a Date field in an Aspose.Cells pivot table (C#)
// AI Prompts: Write C# code that creates a workbook, builds a pivot table from a date column, adds a timeline control linked to that column, sets the timeline’s StartDate to February 1 2023, and configures the level to Month using Aspose.Cells. | Show how to change the timeline granularity from Month to Quarter programmatically in an Aspose.Cells pivot table. | Demonstrate retrieving the Timeline object from a worksheet and updating its StartDate and EndDate to filter the pivot table for any custom date range in C#.
// Common Searches: aspnet c# add timeline control to pivot table for date filtering with Aspose.Cells | set custom start and end dates on Aspose.Cells timeline to filter pivot data | change timeline level to month or quarter in Aspose.Cells .NET example | how to use Aspose.Cells Timeline API to filter sales data by specific months | Aspose.Cells timeline filter example with DateTime values in C#
// Tags: Aspose.Cells add timeline to pivot table C# | timeline startdate filter Aspose.Cells | set timeline level month Aspose.Cells | pivot table date field timeline filter .NET | custom date range timeline Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineFilterDemo
{
    // The sample creates a workbook, fills it with date and sales data, builds a pivot table, adds a timeline linked to the Date field, sets the timeline's start date to February 1 2023, configures the granularity to month, and saves the workbook as TimelineFilterDemo.xlsx.
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

                // Populate sample data with a Date field and a numeric field
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Sales");

                cells["A2"].PutValue(new DateTime(2023, 1, 5));
                cells["B2"].PutValue(1200);

                cells["A3"].PutValue(new DateTime(2023, 2, 12));
                cells["B3"].PutValue(1500);

                cells["A4"].PutValue(new DateTime(2023, 3, 20));
                cells["B4"].PutValue(1800);

                cells["A5"].PutValue(new DateTime(2023, 4, 8));
                cells["B5"].PutValue(2000);

                cells["A6"].PutValue(new DateTime(2023, 5, 15));
                cells["B6"].PutValue(2200);

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B6", "D2", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the Date field to the Row area and the Sales field to the Data area
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot cache and calculate data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline control linked to the Date field of the pivot table
                // The timeline will be placed with its upper‑left corner at cell E2
                sheet.Timelines.Add(pivot, "E2", "Date");

                // Retrieve the created Timeline object
                Timeline timeline = sheet.Timelines[0];

                // Set a custom start date for the timeline (e.g., February 1, 2023)
                timeline.StartDate = new DateTime(2023, 2, 1);

                // Set the timeline granularity to Month
                timeline.CurrentLevel = TimelineLevelType.Month;

                // Save the workbook with the timeline filter applied
                workbook.Save("TimelineFilterDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
