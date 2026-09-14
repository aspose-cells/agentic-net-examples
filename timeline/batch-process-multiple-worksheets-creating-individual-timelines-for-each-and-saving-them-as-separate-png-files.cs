// Title: Batch render each worksheet’s timeline to separate PNG images with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates through all worksheets in a workbook, adds a pivot table and a linked timeline, then saves each sheet as an individual PNG file using Aspose.Cells. | Write a .NET snippet to create a timeline control for every worksheet and export the rendered sheet (including the timeline) to a PNG image. | Provide a C# example that batch processes multiple Excel sheets, attaches a timeline to each pivot table, and uses SheetRender to produce separate PNG files. | Show how to programmatically add a timeline to each worksheet and render each sheet to a PNG file with the one‑page‑per‑sheet option in Aspose.Cells.
// Common Searches: aspnet render timeline from each worksheet to PNG using Aspose.Cells | how to export Excel sheets with timelines as separate images in C# | batch create pivot tables and timelines then save each sheet as PNG Aspose.Cells | C# Aspose.Cells render multiple worksheets to PNG including timeline control | save each worksheet’s timeline as individual PNG file with Aspose.Cells for .NET
// Tags: batch worksheet timeline rendering to PNG with Aspose.Cells | add pivot table and timeline per sheet C# | SheetRender export worksheet as PNG including timeline | one-page-per-sheet image options Aspose.Cells | automated timeline generation for multiple Excel worksheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;

// The example creates a workbook with several worksheets, populates each with date/value data, adds a pivot table and a linked timeline to every sheet, and then uses SheetRender to export each worksheet—including its timeline—to a separate PNG file while also saving the original workbook.
class BatchTimelineRenderer
{
    static void Main()
    {
        try
        {
            // Create a new workbook with at least two worksheets
            Workbook workbook = new Workbook();
            while (workbook.Worksheets.Count < 2)
            {
                workbook.Worksheets.Add();
            }

            // Populate each worksheet with sample date/value data
            for (int i = 0; i < 2; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Name = $"Sheet{i + 1}";

                // Header
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");

                // Sample rows
                for (int row = 0; row < 5; row++)
                {
                    sheet.Cells[row + 1, 0].PutValue(DateTime.Now.AddDays(row));
                    sheet.Cells[row + 1, 1].PutValue(row * 10);
                }
            }

            // Iterate over each worksheet to add a timeline and render to PNG
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Determine the data range (A1:B{lastRow})
                    int lastDataRow = sheet.Cells.MaxDataRow;
                    string dataRange = $"A1:B{lastDataRow + 1}";

                    // Add a pivot table based on the data range, placed at C1
                    int pivotIndex = sheet.PivotTables.Add(dataRange, "C1", "PivotTable");
                    PivotTable pivot = sheet.PivotTables[pivotIndex];

                    // Configure pivot fields: Date as row, Value as data
                    pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                    pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                    // Refresh pivot cache and calculate data
                    pivot.RefreshData();
                    pivot.CalculateData();

                    // Add a timeline linked to the pivot table, positioned at E1
                    sheet.Timelines.Add(pivot, "E1", "Date");

                    // Set image rendering options (PNG, one page per sheet)
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        OnePagePerSheet = true
                    };

                    // Render the worksheet (including the timeline) to a PNG file
                    SheetRender renderer = new SheetRender(sheet, imgOptions);
                    string outputFile = $"{sheet.Name}_Timeline.png";

                    // Ensure the directory for the output file exists
                    string outputDir = Path.GetDirectoryName(outputFile);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    renderer.ToImage(0, outputFile);
                    Console.WriteLine($"Rendered {outputFile}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Save the workbook for reference (optional)
            string workbookPath = "BatchTimelineWorkbook.xlsx";
            try
            {
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved as {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
