// Title: Add a Custom Timeline Filter to a PivotTable Date Field with Aspose.Cells for .NET (C#)
// Description: Demonstrates creating a workbook, inserting Date and Sales data, building a PivotTable, attaching a Timeline control to the Date field, setting a custom start date (Feb 1 2023) and optionally fixing the view to months, then saving the file.
// Keywords: Aspose.Cells | C# | PivotTable | Timeline control | custom start date | date filter | Excel automation | timeline level month | add timeline to pivot | Aspose.Cells example
// Common Searches: Aspose.Cells add timeline to pivot table | C# set custom start date for timeline control | filter PivotTable by date using timeline Aspose.Cells | Aspose.Cells timeline current level month | create interactive date filter in Excel with Aspose.Cells
// Developer Intent: I need to programmatically add a timeline filter to a PivotTable and define its initial date range using Aspose.Cells for .NET.
// Use Cases: Enable end‑users to select custom date ranges in an exported Excel report. | Pre‑configure quarterly sales dashboards with a preset timeline start date. | Automate month‑level filtering for financial data before distribution.
// AI Prompts: Generate C# code to change the timeline view to weeks in Aspose.Cells. | Show how to bind a single timeline to multiple PivotTables in the same workbook. | Explain how to update the selected period of an existing timeline after the workbook is opened.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineFilterDemo
{
    // Demonstrates creating a workbook, inserting Date and Sales data, building a PivotTable, attaching a Timeline control to the Date field, setting a custom start date (Feb 1 2023) and optionally fixing the view to months, then saving the file.
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

                cells["A2"].PutValue(new DateTime(2023, 1, 5));
                cells["B2"].PutValue(1200);

                cells["A3"].PutValue(new DateTime(2023, 2, 12));
                cells["B3"].PutValue(1500);

                cells["A4"].PutValue(new DateTime(2023, 3, 20));
                cells["B4"].PutValue(1800);

                cells["A5"].PutValue(new DateTime(2023, 4, 8));
                cells["B5"].PutValue(2000);

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the Date field to the row area and the Sales field to the data area
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot cache and calculate the pivot table
                pivot.RefreshData();          // Correct API call
                pivot.CalculateData();

                // Add a Timeline control linked to the Date field of the pivot table
                // Place the timeline with its upper‑left corner at cell E1 (row 0, column 4)
                int timelineIndex = sheet.Timelines.Add(pivot, 0, 4, "Date");
                Timeline timeline = sheet.Timelines[timelineIndex];

                // Set a custom start date for the timeline (e.g., February 1, 2023)
                timeline.StartDate = new DateTime(2023, 2, 1);

                // Optionally, set the current level to Month (value 3) so the timeline shows months
                // timeline.CurrentLevel = 3; // Uncomment if you want to enforce month view

                // Save the workbook with the timeline filter applied
                workbook.Save("TimelineFilterDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
