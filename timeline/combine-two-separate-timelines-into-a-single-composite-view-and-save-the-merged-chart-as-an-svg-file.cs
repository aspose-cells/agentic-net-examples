// Title: Combine Pivot Timelines from Two Workbooks and Export as SVG with Aspose.Cells for .NET (C#)
// Description: C# sample that creates two workbooks, each with a pivot table and a date timeline, merges the second workbook into the first using Workbook.Combine, and renders the combined worksheet to a single SVG file. The example shows how to position timelines, avoid overlap, and use SvgImageOptions (FitToViewPort) for a clean composite view.
// Keywords: Aspose.Cells | C# | .NET | combine workbooks | merge timelines | pivot table timeline | export to SVG | SvgImageOptions | FitToViewPort | worksheet rendering | timeline visualization | Excel automation
// Common Searches: Aspose.Cells combine two workbooks C# | merge pivot timelines and export SVG | render Excel timeline as SVG using Aspose.Cells | how to avoid timeline overlap after workbook merge | SvgImageOptions FitToViewPort example
// Developer Intent: Merge two workbooks that each contain a pivot table and a timeline, then generate one SVG image that displays both timelines together.
// Use Cases: Create a unified sales and revenue dashboard by consolidating separate timelines into a single SVG report for web publishing. | Automate quarterly performance reviews by merging data sources and exporting a composite timeline visualization. | Generate a single SVG file with multiple pivot timelines for embedding in presentations or documentation.
// AI Prompts: Write C# code with Aspose.Cells to combine two workbooks, each having a pivot table and timeline, and export the merged sheet as an SVG. | Explain how to configure SvgImageOptions to fit a composite view of multiple timelines to the viewport when rendering to SVG. | Provide best practices for positioning timelines to prevent overlap after merging workbooks before SVG export.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineMergeSvg
{
    // C# sample that creates two workbooks, each with a pivot table and a date timeline, merges the second workbook into the first using Workbook.Combine, and renders the combined worksheet to a single SVG file. The example shows how to position timelines, avoid overlap, and use SvgImageOptions (FitToViewPort) for a clean composite view.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Workbook 1 ----------
                Workbook wb1 = new Workbook();
                Worksheet ws1 = wb1.Worksheets[0];

                // Populate data for first timeline
                ws1.Cells["A1"].Value = "Date";
                ws1.Cells["B1"].Value = "Sales";
                ws1.Cells["A2"].Value = new DateTime(2023, 1, 1);
                ws1.Cells["B2"].Value = 1200;
                ws1.Cells["A3"].Value = new DateTime(2023, 2, 1);
                ws1.Cells["B3"].Value = 1500;
                ws1.Cells["A4"].Value = new DateTime(2023, 3, 1);
                ws1.Cells["B4"].Value = 1800;

                // Create pivot table and timeline for first workbook
                int pivotIdx1 = ws1.PivotTables.Add("A1:B4", "D1", "Pivot1");
                PivotTable pivot1 = ws1.PivotTables[pivotIdx1];
                pivot1.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot1.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivot1.RefreshData();
                pivot1.CalculateData();

                // Place timeline below the data to avoid overlap
                ws1.Timelines.Add(pivot1, 6, 0, "Date");

                // ---------- Workbook 2 ----------
                Workbook wb2 = new Workbook();
                Worksheet ws2 = wb2.Worksheets[0];

                // Populate data for second timeline
                ws2.Cells["A1"].Value = "Date";
                ws2.Cells["B1"].Value = "Revenue";
                ws2.Cells["A2"].Value = new DateTime(2023, 1, 15);
                ws2.Cells["B2"].Value = 800;
                ws2.Cells["A3"].Value = new DateTime(2023, 2, 15);
                ws2.Cells["B3"].Value = 950;
                ws2.Cells["A4"].Value = new DateTime(2023, 3, 15);
                ws2.Cells["B4"].Value = 1100;

                // Create pivot table and timeline for second workbook
                int pivotIdx2 = ws2.PivotTables.Add("A1:B4", "D1", "Pivot2");
                PivotTable pivot2 = ws2.PivotTables[pivotIdx2];
                pivot2.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot2.AddFieldToArea(PivotFieldType.Data, "Revenue");
                pivot2.RefreshData();
                pivot2.CalculateData();

                // Place timeline below the data to avoid overlap
                ws2.Timelines.Add(pivot2, 6, 0, "Date");

                // ---------- Combine ----------
                // Merge the second workbook into the first one
                wb1.Combine(wb2);

                // ---------- Render to SVG ----------
                // Configure SVG rendering options (optional: fit to viewport)
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true
                };

                // Render the first worksheet (which now contains both timelines) to SVG
                SheetRender renderer = new SheetRender(wb1.Worksheets[0], svgOptions);
                renderer.ToImage(0, "CombinedTimelines.svg");

                // The SVG file "CombinedTimelines.svg" now contains the composite view.
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
