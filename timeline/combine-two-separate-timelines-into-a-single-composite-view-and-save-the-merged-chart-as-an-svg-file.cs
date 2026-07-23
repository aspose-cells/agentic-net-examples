// Title: Merge Excel workbooks with timelines and export the combined sheet as SVG using Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to create two workbooks, each with a pivot table and a linked timeline, combine the second workbook into the first, and render the resulting worksheet to an SVG file with FitToViewPort enabled. The merged workbook can also be saved for verification.
// Keywords: Aspose.Cells | C# | .NET | timeline | merge workbooks | combine worksheets | SVG export | SheetRender | SvgImageOptions | pivot table | data consolidation | dashboard visualization
// Common Searches: Aspose.Cells combine workbooks with timelines C# | Export merged Excel sheet to SVG using Aspose.Cells | How to render multiple timelines to a single SVG in .NET | Combine pivot tables from different workbooks and save as SVG | FitToViewPort option for SVG rendering Aspose.Cells
// Developer Intent: Combine two Excel workbooks that each contain a timeline and render the merged worksheet as an SVG image.
// Use Cases: Build a web dashboard that aggregates data from separate Excel sources, each with its own timeline, and displays a single SVG graphic. | Create a printable SVG report that visualizes combined timelines from multiple monthly reports. | Consolidate financial statements with individual timelines into one SVG for presentation or documentation.
// AI Prompts: Generate C# code with Aspose.Cells to merge two workbooks containing timelines and export the combined sheet to SVG. | Explain how timeline connections are preserved when combining workbooks and rendering to SVG with Aspose.Cells. | Recommend SvgImageOptions settings to ensure the merged timeline view fits within the SVG viewport.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace TimelineMergeToSvg
{
    // This example demonstrates how to create two workbooks, each with a pivot table and a linked timeline, combine the second workbook into the first, and render the resulting worksheet to an SVG file with FitToViewPort enabled. The merged workbook can also be saved for verification.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- First workbook with its own timeline ----------
                Workbook wb1 = new Workbook();
                Worksheet ws1 = wb1.Worksheets[0];

                // Sample data for first timeline
                ws1.Cells["A1"].Value = "Date";
                ws1.Cells["B1"].Value = "Value";
                ws1.Cells["A2"].Value = new DateTime(2023, 1, 1);
                ws1.Cells["B2"].Value = 100;
                ws1.Cells["A3"].Value = new DateTime(2023, 2, 1);
                ws1.Cells["B3"].Value = 200;

                // Create pivot table for first timeline
                int pivotIdx1 = ws1.PivotTables.Add("A1:B3", "D1", "Pivot1");
                PivotTable pivot1 = ws1.PivotTables[pivotIdx1];
                pivot1.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot1.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot1.RefreshData();

                // Add timeline linked to the pivot table (using row/column indices)
                // Cell F1 corresponds to row 0, column 5 (zero‑based indices)
                ws1.Timelines.Add(pivot1, 0, 5, "Date");

                // ---------- Second workbook with its own timeline ----------
                Workbook wb2 = new Workbook();
                Worksheet ws2 = wb2.Worksheets[0];

                // Sample data for second timeline
                ws2.Cells["A1"].Value = "Date";
                ws2.Cells["B1"].Value = "Amount";
                ws2.Cells["A2"].Value = new DateTime(2023, 3, 1);
                ws2.Cells["B2"].Value = 150;
                ws2.Cells["A3"].Value = new DateTime(2023, 4, 1);
                ws2.Cells["B3"].Value = 250;

                // Create pivot table for second timeline
                int pivotIdx2 = ws2.PivotTables.Add("A1:B3", "D1", "Pivot2");
                PivotTable pivot2 = ws2.PivotTables[pivotIdx2];
                pivot2.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot2.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot2.RefreshData();

                // Add timeline linked to the second pivot table (cell F1 -> row 0, column 5)
                ws2.Timelines.Add(pivot2, 0, 5, "Date");

                // ---------- Combine the two workbooks ----------
                // The destination workbook (wb1) will receive the content of wb2
                wb1.Combine(wb2);

                // ---------- Render the combined worksheet to SVG ----------
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true
                };
                SheetRender renderer = new SheetRender(wb1.Worksheets[0], svgOptions);
                renderer.ToImage(0, "CombinedTimelines.svg");

                // Optional: save the combined workbook for verification
                wb1.Save("CombinedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
