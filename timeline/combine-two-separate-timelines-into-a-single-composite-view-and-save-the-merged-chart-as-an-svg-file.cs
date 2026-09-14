// Title: Combine two Excel workbooks with pivot table timelines and export the merged sheet as an SVG using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates two workbooks, adds pivot tables with a Date field, attaches timelines, merges the second workbook into the first, and renders the combined worksheet to an SVG file with FitToViewPort using Aspose.Cells. | Show how to rename timeline objects after a workbook merge and then save both the SVG image and the merged workbook in C# with Aspose.Cells. | Provide a step‑by‑step example of accessing the TimelineCollection after combining workbooks and exporting the worksheet that contains multiple timelines to an SVG image.
// Common Searches: asp.net combine two Excel files with pivot table timelines and export to SVG | how to merge timelines from different workbooks using Aspose.Cells C# | render worksheet with multiple timelines to SVG in .NET | Aspose.Cells combine workbooks preserve timelines | C# export combined pivot table timelines as scalable vector graphic
// Tags: combine workbooks with timelines Aspose.Cells | render worksheet to SVG Aspose.Cells | pivot table timeline merging .NET | timeline collection rename Aspose.Cells | export combined timelines as SVG C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineMergeDemo
{
    // The example creates two workbooks, each containing a pivot table and a linked timeline, merges the second workbook into the first, optionally renames the timelines, renders the combined worksheet to an SVG file with FitToViewPort, and saves the merged workbook for verification.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Workbook 1 ----------
                Workbook wb1 = new Workbook();
                Worksheet ws1 = wb1.Worksheets[0];

                // Sample data for first timeline
                ws1.Cells["A1"].Value = "Date";
                ws1.Cells["B1"].Value = "Sales";
                ws1.Cells["A2"].Value = new DateTime(2023, 1, 1);
                ws1.Cells["B2"].Value = 1200;
                ws1.Cells["A3"].Value = new DateTime(2023, 2, 1);
                ws1.Cells["B3"].Value = 1500;

                // Create pivot table
                int ptIdx1 = ws1.PivotTables.Add("A1:B3", "D1", "Pivot1");
                PivotTable pt1 = ws1.PivotTables[ptIdx1];
                pt1.AddFieldToArea(PivotFieldType.Row, "Date");
                pt1.AddFieldToArea(PivotFieldType.Data, "Sales");
                // Refresh pivot cache and calculate data
                pt1.RefreshData();
                pt1.CalculateData();

                // Add timeline linked to the pivot table (placed below the data)
                ws1.Timelines.Add(pt1, 5, 0, "Date");

                // ---------- Workbook 2 ----------
                Workbook wb2 = new Workbook();
                Worksheet ws2 = wb2.Worksheets[0];

                // Sample data for second timeline
                ws2.Cells["A1"].Value = "Date";
                ws2.Cells["B1"].Value = "Revenue";
                ws2.Cells["A2"].Value = new DateTime(2023, 1, 15);
                ws2.Cells["B2"].Value = 800;
                ws2.Cells["A3"].Value = new DateTime(2023, 2, 15);
                ws2.Cells["B3"].Value = 950;

                // Create pivot table
                int ptIdx2 = ws2.PivotTables.Add("A1:B3", "D1", "Pivot2");
                PivotTable pt2 = ws2.PivotTables[ptIdx2];
                pt2.AddFieldToArea(PivotFieldType.Row, "Date");
                pt2.AddFieldToArea(PivotFieldType.Data, "Revenue");
                // Refresh pivot cache and calculate data
                pt2.RefreshData();
                pt2.CalculateData();

                // Add timeline linked to the pivot table (placed below the data)
                ws2.Timelines.Add(pt2, 5, 0, "Date");

                // ---------- Combine Workbooks ----------
                // Merge wb2 into wb1; timelines from both sheets will now reside in wb1
                wb1.Combine(wb2);

                // Optional: Access combined timelines (they are on the first worksheet)
                TimelineCollection timelines = wb1.Worksheets[0].Timelines;
                for (int i = 0; i < timelines.Count; i++)
                {
                    timelines[i].Caption = $"Timeline {i + 1}";
                }

                // ---------- Render to SVG ----------
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true // Ensure the SVG fits the viewport
                };

                // Render the first worksheet (which now contains both timelines) to an SVG file
                SheetRender renderer = new SheetRender(wb1.Worksheets[0], svgOptions);
                renderer.ToImage(0, "CombinedTimelines.svg");

                // Save the combined workbook (optional, for verification)
                wb1.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
