// Title: Create a Dynamic Milestone Timeline with foreach Smart Markers in Aspose.Cells for .NET
// Description: This example builds a List<Milestone>, defines a worksheet template with foreach smart markers (&=Milestones.Date, &=Milestones.Title) in a range named _CellsSmartMarkers, binds the collection to WorkbookDesigner, processes the markers to repeat rows, generates a PivotTable from the populated data, adds an Excel Timeline control linked to the date field, and saves the workbook as MilestonesTimeline.xlsx.
// Keywords: Aspose.Cells | C# | foreach smart marker | smart markers range | timeline control | pivot table | variable length list | Excel timeline | Milestone | WorkbookDesigner | populate collection | GitHub example
// Common Searches: foreach smart marker repeat rows Aspose.Cells | Aspose.Cells create timeline from pivot table | bind List<T> to smart markers C# | generate variable length rows using smart markers | add Excel timeline control with Aspose.Cells
// Developer Intent: Generate a variable‑length list of project milestones and display them on an interactive Excel timeline using foreach smart markers.
// Use Cases: Automatically expand rows for any number of milestones without manual sheet edits. | Build a PivotTable from the generated data to enable timeline filtering. | Integrate an interactive timeline control for project‑tracking dashboards. | Reuse the same template for different projects by supplying a new Milestone collection. | Export the result to an .xlsx file for sharing with stakeholders.
// AI Prompts: Show code to add a Description column to the milestone template and include it in the timeline. | Explain how to set the timeline's start and end dates programmatically. | Provide error handling for null dates or empty titles in the Milestone list. | Demonstrate how to style the timeline control (colors, fonts) using Aspose.Cells. | Convert the generated workbook to PDF while preserving the timeline visual.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsSmartMarkerTimelineDemo
{
    // Simple POCO representing a project milestone
    // This example builds a List<Milestone>, defines a worksheet template with foreach smart markers (&=Milestones.Date, &=Milestones.Title) in a range named _CellsSmartMarkers, binds the collection to WorkbookDesigner, processes the markers to repeat rows, generates a PivotTable from the populated data, adds an Excel Timeline control linked to the date field, and saves the workbook as MilestonesTimeline.xlsx.
    public class Milestone
    {
        public DateTime Date { get; set; }
        public string? Title { get; set; }   // Nullable to satisfy non‑nullable warning
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare sample data (variable‑length list of milestones)
                // -----------------------------------------------------------------
                List<Milestone> milestones = new List<Milestone>
                {
                    new Milestone { Date = new DateTime(2023, 1, 15), Title = "Project Kick‑off" },
                    new Milestone { Date = new DateTime(2023, 2, 10), Title = "Requirement Sign‑off" },
                    new Milestone { Date = new DateTime(2023, 3, 5),  Title = "Design Completion" },
                    new Milestone { Date = new DateTime(2023, 4, 20), Title = "First Prototype" },
                    new Milestone { Date = new DateTime(2023, 6, 30), Title = "Beta Release" },
                    new Milestone { Date = new DateTime(2023, 9, 15), Title = "Final Release" }
                };

                // -----------------------------------------------------------------
                // 2. Create a workbook that will act as the template
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Milestone Date");
                sheet.Cells["B1"].PutValue("Milestone Title");

                // Row that contains foreach smart markers.
                // The markers will be repeated for each item in the "Milestones" collection.
                sheet.Cells["A2"].PutValue("&=Milestones.Date");
                sheet.Cells["B2"].PutValue("&=Milestones.Title");

                // Define the smart‑marker range. Setting the name to "_CellsSmartMarkers"
                // tells the designer to process this range line‑by‑line (i.e., repeat rows).
                Aspose.Cells.Range smartMarkerRange = sheet.Cells.CreateRange("A2:B2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // -----------------------------------------------------------------
                // 3. Process the smart markers using WorkbookDesigner
                // -----------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Bind the data source. The name "Milestones" must match the smart‑marker prefix.
                designer.SetDataSource("Milestones", milestones);
                designer.Process(); // Populate the template with the variable‑length list

                // -----------------------------------------------------------------
                // 4. Create a PivotTable based on the populated data (required for Timeline)
                // -----------------------------------------------------------------
                // Determine the used range after processing
                int lastRow = sheet.Cells.MaxDataRow;
                int lastColumn = sheet.Cells.MaxDataColumn;

                // Add a PivotTable starting at cell D1 (adjust as needed)
                int pivotIndex = sheet.PivotTables.Add(
                    // Source range includes headers and all data rows
                    $"A1:{CellIndexToName(lastRow, lastColumn)}",
                    "D1",
                    "MilestonesPivot");

                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the Date field to the Row area (this will be the timeline base field)
                pivot.AddFieldToArea(PivotFieldType.Row, "Milestone Date");

                // Add the Title field to the Data area (count of milestones per date)
                pivot.AddFieldToArea(PivotFieldType.Data, "Milestone Title");

                // Refresh and calculate the pivot data
                pivot.RefreshData();
                pivot.CalculateData();

                // -----------------------------------------------------------------
                // 5. Add a Timeline control linked to the PivotTable's Date field
                // -----------------------------------------------------------------
                // Use the overload that accepts a destination cell name and the base field name.
                sheet.Timelines.Add(pivot, "F1", "Milestone Date");

                // -----------------------------------------------------------------
                // 6. Save the resulting workbook
                // -----------------------------------------------------------------
                workbook.Save("MilestonesTimeline.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to convert zero‑based row/column indexes to Excel cell name (e.g., 0,0 => "A1")
        private static string CellIndexToName(int rowIndex, int columnIndex)
        {
            // Aspose.Cells provides a utility for this conversion.
            // The CellsHelper class method ColumnIndexToName converts column index to letters.
            // Row index is 0‑based, so add 1 for the Excel row number.
            string columnName = CellsHelper.ColumnIndexToName(columnIndex);
            int rowNumber = rowIndex + 1;
            return $"{columnName}{rowNumber}";
        }
    }
}
