// Title: Create an Excel timeline with Aspose.Cells for .NET and save it as a PNG image
// Common Searches: how to add a timeline control to a pivot table using Aspose.Cells C# | render an Aspose.Cells worksheet that contains a timeline to PNG | export Excel timeline as an image with Aspose.Cells .NET API | fit an Aspose.Cells worksheet to a single page before rendering to image | generate a timeline chart from a date column in Aspose.Cells and save as PNG
// Tags: Aspose.Cells timeline PNG rendering | C# timeline control integration with pivot table | Aspose.Cells export worksheet as PNG | fit worksheet to single page Aspose.Cells | create pivot table with date field Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The example creates a workbook, fills it with ship dates and sales, builds a pivot table, adds a timeline control for the date field, configures the page to fit on one page, and renders the worksheet (including the timeline) to a PNG file while also saving the workbook as an XLSX file.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Output image file
            string outputImagePath = "timeline.png";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // -------------------------------------------------
            // 1. Populate sample data (Ship Date and Sales columns)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Ship Date");
            sheet.Cells["B1"].PutValue("Sales");

            DateTime[] dates = new DateTime[]
            {
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 12),
                new DateTime(2023, 1, 20),
                new DateTime(2023, 2, 3),
                new DateTime(2023, 2, 15),
                new DateTime(2023, 3, 1)
            };

            double[] sales = new double[] { 1500, 2300, 1800, 2100, 2500, 3000 };

            for (int i = 0; i < dates.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(dates[i]);
                sheet.Cells[i + 1, 1].PutValue(sales[i]);
            }

            // -------------------------------------------------
            // 2. Create a PivotTable based on the data range
            // -------------------------------------------------
            string sourceData = $"A1:B{dates.Length + 1}";

            // Add pivot table at cell D3 with a default name
            int pivotIndex = sheet.PivotTables.Add(sourceData, "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Ship Date");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // -------------------------------------------------
            // 3. Add a Timeline control for the date field
            // -------------------------------------------------
            // Place the timeline starting at G3 (row index 2, column index 6)
            sheet.Timelines.Add(pivotTable, 2, 6, "Ship Date");

            // -------------------------------------------------
            // 4. Adjust page setup so the timeline fits on one page
            // -------------------------------------------------
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.SetFitToPages(1, 1); // fit entire sheet to a single page

            // -------------------------------------------------
            // 5. Render the worksheet (including the timeline) to PNG
            // -------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png
            };

            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            if (sheetRender.PageCount > 0)
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputImagePath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                sheetRender.ToImage(0, outputImagePath);
                Console.WriteLine($"Timeline image saved successfully to '{Path.GetFullPath(outputImagePath)}'.");
            }
            else
            {
                Console.WriteLine("No pages were generated for rendering. Check page setup and data.");
            }

            // Optional: Save the workbook for verification
            string workbookPath = "TimelineDemo.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(workbookPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the timeline image:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
