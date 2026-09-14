// Title: How to render an Aspose.Cells Timeline to PNG and embed it as a centered header image with a title in a PDF report (C#)
// AI Prompts: Generate C# code that creates a pivot table, adds a Timeline linked to the Date field, renders the worksheet to a PNG image, and sets that image as the center header picture while adding a left-aligned title before saving as PDF. | Show how to adjust the size of a rendered Timeline image and embed it into the PDF header using Aspose.Cells PageSetup methods. | Explain the steps to export a workbook with a custom header that contains both text and a timeline graphic, including memory‑stream handling for the PNG.
// Common Searches: Aspose.Cells C# render timeline control to image for PDF header | Set header picture and text in PDF export using Aspose.Cells .NET | How to add a timeline graphic to the header of a PDF report with Aspose.Cells | C# Aspose.Cells render worksheet as PNG and use it in page header | Embedding pivot table timeline image in PDF header Aspose.Cells example
// Tags: render timeline to PNG Aspose.Cells | set header picture PDF Aspose.Cells | center header image with title Aspose.Cells | timeline image in PDF header C# | pivot table timeline export PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;

// Demonstrates creating a workbook with sample data, building a pivot table, adding a Timeline control, rendering the sheet to a PNG image in memory, inserting the image as a centered header picture while adding a left-aligned title, and saving the workbook as a PDF using Aspose.Cells for .NET.
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
            sheet.Cells["A1"].Value = "Date";
            sheet.Cells["B1"].Value = "Value";
            sheet.Cells["A2"].Value = new DateTime(2023, 1, 1);
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = new DateTime(2023, 2, 1);
            sheet.Cells["B3"].Value = 150;
            sheet.Cells["A4"].Value = new DateTime(2023, 3, 1);
            sheet.Cells["B4"].Value = 180;

            // 3. Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh pivot cache and calculate data
            pivot.RefreshData();
            pivot.CalculateData();

            // 4. Add a Timeline control linked to the pivot table (using the "Date" field)
            int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "Date");
            // Optional size adjustments:
            // sheet.Timelines[timelineIdx].WidthPixel = 400;
            // sheet.Timelines[timelineIdx].HeightPixel = 80;

            // 5. Render the worksheet (which now contains the timeline) to an image in memory
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            byte[] timelineImageBytes;
            using (MemoryStream imgStream = new MemoryStream())
            {
                renderer.ToImage(0, imgStream); // render first (and only) page
                timelineImageBytes = imgStream.ToArray();
            }

            // 6. Insert the rendered timeline image into the header (center section)
            sheet.PageSetup.SetHeaderPicture(1, timelineImageBytes); // 1 = center header
            sheet.PageSetup.SetHeader(0, "Report Title"); // left header
            sheet.PageSetup.SetHeader(1, "&G");          // center header displays the picture

            // 7. Save the workbook as a PDF; the header (title + timeline image) will appear on each page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DisplayDocTitle = true
            };
            workbook.Save("TimelineReport.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
