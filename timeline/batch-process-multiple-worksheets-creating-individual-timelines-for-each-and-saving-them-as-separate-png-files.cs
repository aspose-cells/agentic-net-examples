using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineBatch
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Number of worksheets to process
                int sheetCount = 3;

                // Loop to create worksheets with sample data, pivot tables, timelines and render to PNG
                for (int i = 0; i < sheetCount; i++)
                {
                    // Add a new worksheet (first iteration uses the default sheet)
                    Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                    // Populate sample data: Date column and Value column
                    sheet.Cells["A1"].PutValue("Date");
                    sheet.Cells["B1"].PutValue("Value");
                    for (int row = 2; row <= 6; row++)
                    {
                        sheet.Cells[$"A{row}"].PutValue(DateTime.Today.AddDays(row - 2));
                        sheet.Cells[$"B{row}"].PutValue(row * 10);
                    }

                    // Create a pivot table based on the sample data
                    int pivotIndex = sheet.PivotTables.Add("A1:B6", "D1", $"Pivot{i + 1}");
                    PivotTable pivot = sheet.PivotTables[pivotIndex];
                    pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                    pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                    pivot.RefreshData();
                    pivot.CalculateData();

                    // Add a timeline linked to the pivot table (placed at cell F1)
                    sheet.Timelines.Add(pivot, "F1", "Date");

                    // Configure image rendering options for PNG output
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        // Image type is inferred from the file extension, so we only set pagination
                        OnePagePerSheet = true
                    };

                    // Render the worksheet (including the timeline) to a PNG file
                    SheetRender renderer = new SheetRender(sheet, imgOptions);
                    string outputPath = Path.Combine(Environment.CurrentDirectory, $"Sheet{i + 1}_Timeline.png");
                    renderer.ToImage(0, outputPath);
                    renderer.Dispose();

                    Console.WriteLine($"Rendered {sheet.Name} with timeline to: {outputPath}");
                }

                // Save the workbook for reference (optional)
                string workbookPath = Path.Combine(Environment.CurrentDirectory, "TimelinesWorkbook.xlsx");
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}