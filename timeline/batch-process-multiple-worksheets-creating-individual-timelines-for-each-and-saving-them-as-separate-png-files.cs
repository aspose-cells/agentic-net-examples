using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;

namespace AsposeCellsBatchTimeline
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Number of worksheets to process
                int sheetCount = 3;

                // Prepare sample data for each worksheet and add a timeline
                for (int i = 0; i < sheetCount; i++)
                {
                    // Add a new worksheet (first one already exists)
                    Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                    // Populate sample data: Date column and Sales column
                    sheet.Cells["A1"].PutValue("Date");
                    sheet.Cells["B1"].PutValue("Sales");
                    for (int row = 2; row <= 6; row++)
                    {
                        sheet.Cells[$"A{row}"].PutValue(new DateTime(2023, 1, row - 1));
                        sheet.Cells[$"B{row}"].PutValue(100 * row);
                    }

                    // Create a pivot table using the sample data
                    int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", $"Pivot{i + 1}");
                    PivotTable pivot = sheet.PivotTables[pivotIndex];

                    // Add the Date field to the Page area (required for Timeline)
                    pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                    // Add the Sales field to the Data area
                    pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                    // Refresh and calculate the pivot table
                    pivot.RefreshData();
                    pivot.CalculateData();

                    // Add a timeline linked to the pivot table
                    sheet.Timelines.Add(pivot, "F1", "Date");
                }

                // Render each worksheet (with its timeline) to a separate PNG file
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Configure image rendering options
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        ImageType = Aspose.Cells.Drawing.ImageType.Png,
                        OnePagePerSheet = true
                    };

                    // Create a SheetRender for the current worksheet
                    SheetRender renderer = new SheetRender(sheet, imgOptions);

                    // Render the first page of the sheet to a PNG file
                    string outputPath = Path.Combine(Environment.CurrentDirectory, $"{sheet.Name}_Timeline.png");
                    renderer.ToImage(0, outputPath);
                }

                // Optional: Save the workbook for reference
                string workbookPath = Path.Combine(Environment.CurrentDirectory, "BatchTimelineWorkbook.xlsx");
                workbook.Save(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}