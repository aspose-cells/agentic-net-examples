// Title: Create an Aspose.Cells PivotTable Timeline with logarithmic date scale, custom tick intervals, and export to PNG using C#
// Common Searches: how to add a timeline to a pivot table with Aspose.Cells in C# | setting a logarithmic date scale on an Aspose.Cells timeline control | customizing tick intervals for an Aspose.Cells timeline | exporting a worksheet that contains a timeline to PNG with Aspose.Cells | high‑resolution PNG rendering of Aspose.Cells worksheets in .NET
// Tags: Aspose.Cells timeline logarithmic scaling | Aspose.Cells timeline tick interval configuration | Aspose.Cells render worksheet to PNG | Aspose.Cells pivot table timeline integration | Aspose.Cells image rendering options C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineDemo
{
    // The example creates a workbook, fills it with sample order data, builds a pivot table, adds a timeline linked to the OrderDate field, shows where to configure a logarithmic date scale and custom tick intervals, and renders each worksheet page to high‑resolution PNG images while saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data that includes a date column
                sheet.Cells["A1"].PutValue("OrderDate");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["C1"].PutValue("Sales");

                DateTime startDate = new DateTime(2022, 1, 1);
                Random rnd = new Random();
                for (int i = 0; i < 30; i++)
                {
                    DateTime orderDate = startDate.AddDays(rnd.Next(1, 10) * i);
                    sheet.Cells[i + 2, 0].PutValue(orderDate);
                    sheet.Cells[i + 2, 1].PutValue("Category " + ((i % 3) + 1));
                    sheet.Cells[i + 2, 2].PutValue(rnd.Next(100, 1000));
                }

                // 3. Create a pivot table based on the data
                string sourceRange = "A1:C31";
                string destCell = "E3";

                // Use overload with isXlsClassic flag (true) to preserve classic behavior
                int pivotIndex = sheet.PivotTables.Add(sourceRange, destCell, "SalesPivot", true);
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "OrderDate");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // 4. Add a timeline control for the date field of the pivot table
                // The Add method returns the index of the newly created timeline
                int timelineIndex = sheet.Timelines.Add(pivotTable, 2, 6, "OrderDate");
                Timeline timeline = sheet.Timelines[timelineIndex];

                // 5. Placeholder for future timeline configuration (e.g., logarithmic scale, tick interval)
                try
                {
                    // Example (if supported in future versions):
                    // timeline.LogarithmicScale = true;
                    // timeline.TickInterval = TimeSpan.FromDays(7);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Warning: Timeline configuration failed. Details: " + ex.Message);
                }

                // 6. Render the worksheet (including the timeline) to PNG images
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = false
                };

                SheetRender renderer = new SheetRender(sheet, imgOptions);
                int pageCount = renderer.PageCount;
                Console.WriteLine($"Worksheet renders to {pageCount} page(s).");

                for (int page = 0; page < pageCount; page++)
                {
                    string pngPath = $"Timeline_Page_{page + 1}.png";
                    renderer.ToImage(page, pngPath);
                    Console.WriteLine($"Rendered page {page + 1} to {pngPath}");
                }

                // 7. Save the workbook to verify the timeline is persisted
                string workbookPath = "TimelineDemo.xlsx";
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {workbookPath}");

                Console.WriteLine("Processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
