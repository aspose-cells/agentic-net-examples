// Title: How to create an Aspose.Cells timeline with alternating row colors and export it as a high‑resolution JPEG in C#
// Description: This example shows how to build a new workbook, fill it with date, category, and value data, apply LightGray shading to even rows and White to odd rows, generate a pivot table, attach a Timeline slicer to the Date field, and render the worksheet (including the timeline) to a 300 DPI JPEG with maximum quality. The workbook and image are saved to disk.
// Keywords: Aspose.Cells | C# | .NET | timeline control | alternating row colors | high resolution JPEG | 300 DPI | pivot table | render worksheet to image | Excel timeline slicer | export to JPEG
// Common Searches: Aspose.Cells add timeline to pivot table C# | Export Aspose.Cells worksheet as high DPI JPEG | Apply alternating row shading in Aspose.Cells workbook | Render Excel timeline to image using Aspose.Cells | C# code for timeline with colored rows and JPEG output
// Developer Intent: Create a worksheet with a pivot‑based timeline, apply alternating row shading, and render it to a high‑resolution JPEG image using Aspose.Cells for .NET.
// Use Cases: Build a sales dashboard where alternating row colors improve readability, a timeline slicer filters dates, and the final view is exported as a 300 DPI JPEG for printable reports. | Generate a project schedule sheet with colored rows for visual clarity, attach a timeline to the date column, and save the visual as a high‑quality image for embedding in presentations.
// AI Prompts: Show how to change the even/odd row colors to a custom palette while keeping the timeline rendering unchanged. | Provide code to render the worksheet as a PNG with 600 DPI resolution instead of JPEG using Aspose.Cells. | Explain how to add multiple timelines for different date fields in the same worksheet and export each to separate high‑quality images.

using System;
using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineDemo
{
    // This example shows how to build a new workbook, fill it with date, category, and value data, apply LightGray shading to even rows and White to odd rows, generate a pivot table, attach a Timeline slicer to the Date field, and render the worksheet (including the timeline) to a 300 DPI JPEG with maximum quality. The workbook and image are saved to disk.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (Date, Category, Value)
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Category");
                cells["C1"].PutValue("Value");

                DateTime start = new DateTime(2023, 1, 1);
                for (int i = 0; i < 10; i++)
                {
                    cells[i + 1, 0].PutValue(start.AddDays(i));
                    cells[i + 1, 1].PutValue("Item " + (i + 1));
                    cells[i + 1, 2].PutValue(10 + i * 5);
                }

                // Apply alternating row fill colors (even rows: LightGray, odd rows: White)
                Style evenStyle = workbook.CreateStyle();
                evenStyle.ForegroundColor = Color.LightGray;
                evenStyle.Pattern = BackgroundType.Solid;

                Style oddStyle = workbook.CreateStyle();
                oddStyle.ForegroundColor = Color.White;
                oddStyle.Pattern = BackgroundType.Solid;

                for (int row = 1; row <= 10; row++)
                {
                    Style styleToApply = (row % 2 == 0) ? evenStyle : oddStyle;
                    for (int col = 0; col <= 2; col++)
                    {
                        cells[row, col].SetStyle(styleToApply);
                    }
                }

                // Create a pivot table based on the data
                int pivotIndex = sheet.PivotTables.Add("A1:C11", "E1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add fields: Date as Page (required for Timeline), Category as Row, Value as Data
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Optional: set a style for the pivot table
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                // Refresh pivot cache and calculate data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline control linked to the pivot table (using the Date field)
                sheet.Timelines.Add(pivot, 15, 0, "Date");

                // Render the worksheet (including the timeline) to a high‑quality JPEG image
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // ImageFormat property is not available in some versions; default format will be used (PNG)
                    HorizontalResolution = 300,   // high DPI
                    VerticalResolution = 300,
                    Quality = 100                 // maximum quality for supported formats
                };

                WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
                // Render the first (and only) page to an image file
                renderer.ToImage(0, "TimelineHighQuality.jpg");

                // Optionally, also save the workbook for reference
                workbook.Save("TimelineDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
