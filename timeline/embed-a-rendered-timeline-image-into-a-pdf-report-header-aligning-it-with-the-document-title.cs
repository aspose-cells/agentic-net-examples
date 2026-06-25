using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

class TimelineHeaderPdfReport
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
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(300);

            // 3. Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // 4. Add a timeline control linked to the pivot table (centered on column B)
            // Note: Timeline class may not be available in older Aspose.Cells versions,
            // so we only add the timeline and skip direct property manipulation.
            int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "Date");
            // Optional: adjust size/position if supported by the library version
            // sheet.Timelines[timelineIdx].WidthPixel = 400;
            // sheet.Timelines[timelineIdx].HeightPixel = 80;

            // 5. Render the worksheet (which now contains the timeline) to a PNG image in memory
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            byte[] timelineImage;
            using (MemoryStream imgStream = new MemoryStream())
            {
                renderer.ToImage(0, imgStream);          // render first page
                timelineImage = imgStream.ToArray();     // capture image bytes
            }

            // 6. Set the rendered timeline image as the center header picture (section 1 = center)
            sheet.PageSetup.SetHeaderPicture(1, timelineImage);

            // 7. Define the header script: display the picture (&G) followed by the report title
            sheet.PageSetup.SetHeader(1, "&G My Report Title");

            // 8. Prepare PDF save options to show the document title in the PDF viewer title bar
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DisplayDocTitle = true
            };

            // 9. Save the workbook as a PDF; the header now contains the timeline image aligned with the title
            string outputPath = "TimelineHeaderReport.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}