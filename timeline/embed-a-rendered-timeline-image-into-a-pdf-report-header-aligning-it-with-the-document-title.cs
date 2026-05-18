using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTimelinePdfHeader
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data (date and value) for the pivot table
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
                sheet.Cells["B4"].PutValue(180);

                // 3. Create a pivot table based on the data
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.RefreshData();
                pivot.CalculateData();

                // 4. Add a Timeline control linked to the pivot table (using the "Date" field)
                int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "Date");
                // Optional: adjust size/position if needed (requires Timeline type support)
                // sheet.Timelines[timelineIdx].WidthPixel = 400;
                // sheet.Timelines[timelineIdx].HeightPixel = 80;

                // 5. Render the worksheet (which now contains the timeline) to an image in memory
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true
                };
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                using (MemoryStream imgStream = new MemoryStream())
                {
                    // Render first (and only) page to the stream
                    renderer.ToImage(0, imgStream);
                    byte[] imgBytes = imgStream.ToArray();

                    // 6. Set the rendered timeline image as a header picture (center section)
                    // Section index: 1 = center
                    sheet.PageSetup.SetHeaderPicture(1, imgBytes);

                    // 7. Build the header script: title text followed by the image placeholder (&G)
                    sheet.PageSetup.SetHeader(1, "Report Title &G");
                }

                // 8. Save the workbook as PDF, ensuring header/footer is included
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    DisplayDocTitle = true
                };
                string outputPath = "TimelineReport.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}