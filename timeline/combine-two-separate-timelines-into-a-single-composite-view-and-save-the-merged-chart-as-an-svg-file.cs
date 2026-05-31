using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace TimelineMergeToSvg
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create first workbook with a timeline ----------
                Workbook wb1 = new Workbook();
                Worksheet ws1 = wb1.Worksheets[0];

                // Sample data
                ws1.Cells["A1"].Value = "Date";
                ws1.Cells["B1"].Value = "Sales";
                ws1.Cells["A2"].Value = new DateTime(2023, 1, 1);
                ws1.Cells["B2"].Value = 1200;
                ws1.Cells["A3"].Value = new DateTime(2023, 2, 1);
                ws1.Cells["B3"].Value = 1500;
                ws1.Cells["A4"].Value = new DateTime(2023, 3, 1);
                ws1.Cells["B4"].Value = 1800;

                // Pivot table for timeline source
                int ptIdx1 = ws1.PivotTables.Add("A1:B4", "D1", "Pivot1");
                PivotTable pt1 = ws1.PivotTables[ptIdx1];
                pt1.AddFieldToArea(PivotFieldType.Row, "Date");
                pt1.AddFieldToArea(PivotFieldType.Data, "Sales");
                pt1.RefreshData();
                pt1.CalculateData();

                // Add timeline linked to the pivot table (place it below the data to avoid overlap)
                ws1.Timelines.Add(pt1, 5, 0, "Date");

                // ---------- Create second workbook with a timeline ----------
                Workbook wb2 = new Workbook();
                Worksheet ws2 = wb2.Worksheets[0];

                // Sample data
                ws2.Cells["A1"].Value = "Date";
                ws2.Cells["B1"].Value = "Revenue";
                ws2.Cells["A2"].Value = new DateTime(2023, 1, 15);
                ws2.Cells["B2"].Value = 800;
                ws2.Cells["A3"].Value = new DateTime(2023, 2, 15);
                ws2.Cells["B3"].Value = 950;
                ws2.Cells["A4"].Value = new DateTime(2023, 3, 15);
                ws2.Cells["B4"].Value = 1100;

                // Pivot table for timeline source
                int ptIdx2 = ws2.PivotTables.Add("A1:B4", "D1", "Pivot2");
                PivotTable pt2 = ws2.PivotTables[ptIdx2];
                pt2.AddFieldToArea(PivotFieldType.Row, "Date");
                pt2.AddFieldToArea(PivotFieldType.Data, "Revenue");
                pt2.RefreshData();
                pt2.CalculateData();

                // Add timeline linked to the second pivot table (place it below the data)
                ws2.Timelines.Add(pt2, 5, 0, "Date");

                // ---------- Combine the two workbooks ----------
                // The destination workbook (wb1) will receive the content of wb2
                wb1.Combine(wb2);

                // ---------- Render the combined worksheet to SVG ----------
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true // make SVG fit the viewport
                };

                // Render the first worksheet (index 0) to SVG
                SheetRender renderer = new SheetRender(wb1.Worksheets[0], svgOptions);
                renderer.ToImage(0, "CombinedTimelines.svg");

                Console.WriteLine("Combined timelines rendered to SVG successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}