// Title: Aspose.Cells for .NET – Move PivotTable Slicer to Cell D5 and Align Height with Chart
// Description: This C# example shows how to create a workbook, add sample data, build a pivot table, insert a column chart, and add a slicer for the pivot's Category field. It then moves the slicer so its upper‑left corner aligns with cell D5 (UpperLeftColumn = 3, UpperLeftRow = 4) and sets the slicer height to match the chart height before saving the file.
// Keywords: Aspose.Cells | C# slicer positioning | move slicer to D5 | slicer height alignment | pivot table slicer .NET | chart synchronization | Excel slicer placement | UpperLeftColumn UpperLeftRow | Aspose.Cells example | programmatic slicer move
// Common Searches: Aspose.Cells move slicer to specific cell | Set slicer UpperLeftColumn UpperLeftRow .NET | Align slicer height with chart Aspose.Cells | C# code to reposition pivot slicer | Place Excel slicer at D5 using Aspose.Cells
// Developer Intent: Reposition a pivot slicer to cell D5 and match its height to an existing chart.
// Use Cases: Designing automated Excel dashboards where slicers must line up with charts | Generating reports with consistent UI by programmatically aligning slicer and chart dimensions | Batch‑creating worksheets that share a common layout of slicers and charts
// AI Prompts: Write C# code with Aspose.Cells to move a pivot slicer to cell D5 and set its height equal to a column chart. | Explain how UpperLeftColumn and UpperLeftRow map to Excel cell addresses in Aspose.Cells. | Provide a reusable method that aligns multiple slicers with their corresponding charts in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// This C# example shows how to create a workbook, add sample data, build a pivot table, insert a column chart, and add a slicer for the pivot's Category field. It then moves the slicer so its upper‑left corner aligns with cell D5 (UpperLeftColumn = 3, UpperLeftRow = 4) and sets the slicer height to match the chart height before saving the file.
class MoveSlicerExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B1"].Value = "Sales";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["B3"].Value = 150;
            sheet.Cells["B4"].Value = 200;

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a column chart that visualizes the pivot data
            int chartIdx = sheet.Charts.Add(ChartType.Column, 2, 0, 12, 5);
            Chart chart = sheet.Charts[chartIdx];

            // Build the data range for the series (values column of the pivot)
            string dataRange = $"{CellsHelper.CellIndexToName(pivot.TableRange1.StartRow + 1, pivot.TableRange1.StartColumn + 1)}:" +
                               $"{CellsHelper.CellIndexToName(pivot.TableRange1.EndRow, pivot.TableRange1.StartColumn + 1)}";
            chart.NSeries.Add(dataRange, true);

            // Build the category range (row labels column of the pivot)
            string categoryRange = $"{CellsHelper.CellIndexToName(pivot.TableRange1.StartRow + 1, pivot.TableRange1.StartColumn)}:" +
                                   $"{CellsHelper.CellIndexToName(pivot.TableRange1.EndRow, pivot.TableRange1.StartColumn)}";
            chart.NSeries.CategoryData = categoryRange;

            // Add a slicer for the pivot table (initially placed at A6)
            int slicerIdx = sheet.Slicers.Add(pivot, "A6", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Move the slicer so its upper‑left corner aligns with cell D5 (zero‑based indices)
            slicer.Shape.UpperLeftColumn = 3; // Column D
            slicer.Shape.UpperLeftRow = 4;    // Row 5

            // Align the slicer height with the chart height
            slicer.Shape.Height = chart.ChartObject.Height;

            // Save the workbook
            workbook.Save("SlicerMoved.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
