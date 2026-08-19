// Title: Batch render worksheet timelines to PNG with Aspose.Cells for .NET
// Description: Creates a workbook with several worksheets, adds sample data, builds a pivot table on each sheet, attaches a timeline to the Date field, and renders every worksheet (including its timeline) to an individual PNG file using Aspose.Cells. The full workbook is also saved as an XLSX for reference.
// Keywords: Aspose.Cells C# | timeline rendering | pivot table timeline PNG | batch export worksheets | SheetRender PNG | ImageOrPrintOptions one page per sheet | export timeline image | multiple worksheets Aspose.Cells | save workbook with timelines | C# render timeline to image
// Common Searches: Aspose.Cells render timeline to PNG C# | Batch export worksheet timelines as images | How to save each sheet as PNG with timeline Aspose.Cells | One page per sheet image rendering Aspose.Cells | Export pivot table timeline picture .NET
// Developer Intent: Generate a timeline on every worksheet and export each sheet as a separate PNG image.
// Use Cases: Create regional sales dashboards where each region's timeline is delivered as a PNG snapshot. | Automate project phase visualizations across multiple sheets and publish the images to a web portal. | Produce a master workbook with embedded timelines while providing individual PNG files for reporting tools.
// AI Prompts: Write C# code that adds a timeline to each worksheet's pivot table and renders each sheet to a PNG using Aspose.Cells. | Explain how to configure ImageOrPrintOptions for one‑page‑per‑sheet rendering of timelines in Aspose.Cells. | Refactor the rendering loop to accept a variable number of worksheets and a custom output directory.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;

// Creates a workbook with several worksheets, adds sample data, builds a pivot table on each sheet, attaches a timeline to the Date field, and renders every worksheet (including its timeline) to an individual PNG file using Aspose.Cells. The full workbook is also saved as an XLSX for reference.
class BatchTimelineRenderer
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Number of worksheets to generate
            int numberOfSheets = 3;

            // -------------------------------------------------------------
            // 1. Create worksheets, add sample data, pivot tables and timelines
            // -------------------------------------------------------------
            for (int i = 0; i < numberOfSheets; i++)
            {
                // Use the first sheet for i == 0, otherwise add new sheets
                Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                // Populate sample data (Date and Value columns)
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");
                for (int row = 2; row <= 6; row++)
                {
                    sheet.Cells[row - 1, 0].PutValue(DateTime.Today.AddDays(row - 2));
                    sheet.Cells[row - 1, 1].PutValue(row * 10);
                }

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", $"Pivot{i + 1}");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add Date as a page (filter) field – required for timeline
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                // Add Date as a row field (optional, for display)
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                // Add Value as a data field
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh pivot data and calculate results
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline linked to the pivot table (placed at cell F1)
                sheet.Timelines.Add(pivot, "F1", "Date");
            }

            // ---------------------------------------------------------------
            // 2. Render each worksheet (with its timeline) to a separate PNG
            // ---------------------------------------------------------------
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Rendering options: one page per sheet
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                };

                // Render the worksheet to an image; format inferred from file extension
                SheetRender renderer = new SheetRender(sheet, renderOptions);
                string outputFile = $"Timeline_Sheet{i + 1}.png";
                renderer.ToImage(0, outputFile);
            }

            // Save the workbook containing all timelines for reference
            workbook.Save("WorkbookWithTimelines.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
