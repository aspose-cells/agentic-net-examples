// Title: Apply a dark workbook theme, customize timeline appearance, and export an Aspose.Cells workbook to PDF in C#
// AI Prompts: Generate C# code that creates a pivot table, adds a linked timeline, sets dark theme colors (Background1, Text1, Accent1, Accent2), hides the timeline header, and saves the workbook as a PDF with a custom gridline color. | Show how to change the timeline shape's line weight and border color using Aspose.Cells before exporting the workbook to PDF. | Demonstrate configuring ThemeColorType values for a dark theme and setting PdfSaveOptions properties in Aspose.Cells C#.
// Common Searches: c# aspose.cells how to set dark theme colors for a workbook | aspose.cells timeline hide header and change border thickness | export workbook with timeline to pdf using aspose.cells c# | set pdf gridline color when saving aspose.cells workbook | apply dark theme to timeline control in aspose.cells example
// Tags: set workbook dark theme Aspose.Cells | customize timeline border Aspose.Cells C# | hide timeline header Aspose.Cells | export timeline to PDF Aspose.Cells | configure PdfSaveOptions gridline color Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineDarkTheme
{
    // The example creates a workbook with sample data, builds a pivot table, adds a linked timeline, applies a dark theme by setting specific ThemeColorType values, hides the timeline header, adjusts the timeline border thickness, and saves the workbook as a PDF with a defined gridline color.
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

                // Populate sample data (date and value) for the pivot table
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue(new DateTime(2023, 1, 1));
                cells["B2"].PutValue(120);
                cells["A3"].PutValue(new DateTime(2023, 2, 1));
                cells["B3"].PutValue(150);
                cells["A4"].PutValue(new DateTime(2023, 3, 1));
                cells["B4"].PutValue(180);
                cells["A5"].PutValue(new DateTime(2023, 4, 1));
                cells["B5"].PutValue(200);

                // Create a pivot table based on the data range
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh pivot cache
                pivot.RefreshData();

                // Add a Timeline linked to the pivot table (placed at row 10, column 1)
                int timelineIdx = sheet.Timelines.Add(pivot, 10, 1, "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];

                // Apply a dark theme to the workbook (background dark, text light)
                workbook.SetThemeColor(ThemeColorType.Background1, Color.FromArgb(30, 30, 30));
                workbook.SetThemeColor(ThemeColorType.Text1, Color.FromArgb(220, 220, 220));
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(70, 130, 180)); // steel blue
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 165, 0)); // orange

                // Adjust the Timeline's border appearance
                timeline.Shape.Line.Weight = 1.5; // thicker line

                // Optionally hide the header to emphasize the dark look
                timeline.ShowHeader = false;

                // Prepare PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    GridlineColor = Color.Black
                };

                // Save the workbook (including the timeline) as a PDF
                workbook.Save("TimelineDarkTheme.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
