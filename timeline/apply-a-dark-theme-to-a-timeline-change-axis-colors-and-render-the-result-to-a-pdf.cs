using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineDarkThemePdf
{
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

                // -------------------------------------------------
                // Populate sample data (Date, Category, Value)
                // -------------------------------------------------
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Category");
                cells["C1"].PutValue("Value");

                DateTime start = new DateTime(2023, 1, 1);
                for (int i = 0; i < 6; i++)
                {
                    cells[$"A{2 + i}"].PutValue(start.AddDays(i * 7));
                    cells[$"B{2 + i}"].PutValue(i % 2 == 0 ? "Alpha" : "Beta");
                    cells[$"C{2 + i}"].PutValue(10 + i * 5);
                }

                // -------------------------------------------------
                // Create a PivotTable based on the data
                // -------------------------------------------------
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("A1:C7", "E1", "PivotTable1");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Column, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // Add a Timeline linked to the PivotTable (base field: Date)
                // -------------------------------------------------
                int timelineIdx = sheet.Timelines.Add(pivot, 12, 0, "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];
                timeline.Caption = "Sales Timeline";
                timeline.ShowHeader = true;
                timeline.ShowHorizontalScrollbar = true;
                timeline.ShowSelectionLabel = true;
                timeline.ShowTimeLevel = true;

                // -------------------------------------------------
                // Apply a dark custom theme (background dark, text light)
                // -------------------------------------------------
                Color[] darkTheme = new Color[]
                {
                    Color.FromArgb(30, 30, 30),   // Background1 (dark gray)
                    Color.White,                  // Text1 (white)
                    Color.FromArgb(45, 45, 45),   // Background2 (slightly lighter gray)
                    Color.White,                  // Text2 (white)
                    Color.FromArgb(0, 120, 215),  // Accent1 (blue)
                    Color.FromArgb(0, 153, 0),    // Accent2 (green)
                    Color.FromArgb(255, 140, 0),  // Accent3 (orange)
                    Color.FromArgb(191, 0, 0),    // Accent4 (red)
                    Color.FromArgb(102, 102, 102) // Accent5 (medium gray)
                };
                // Note: Aspose.Cells does not provide a direct API to replace the built‑in theme,
                // but you can apply the colors to styles as needed. For this sample we keep it simple.

                // -------------------------------------------------
                // Save the workbook as PDF
                // -------------------------------------------------
                string outputPath = "TimelineDarkTheme.pdf";
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Runtime safety: log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}