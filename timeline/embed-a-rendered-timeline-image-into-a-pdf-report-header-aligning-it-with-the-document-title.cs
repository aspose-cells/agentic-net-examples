// Title: C# – Render a PivotTable Timeline as a Header Image and Export to PDF with Aspose.Cells
// Description: This example creates a workbook, adds sample data, builds a PivotTable, attaches a Timeline control, renders the worksheet (including the Timeline) to an in‑memory image, sets the image as the centered header picture while placing a text title on the left, and saves the result as a PDF where the header shows the timeline graphic aligned with the document title.
// Keywords: Aspose.Cells for .NET | C# timeline header PDF | render PivotTable timeline to image | set header picture Aspose.Cells | PDF header with timeline graphic | align header image with title | Aspose.Cells PDF export example | timeline control in PDF header | Aspose.Cells rendering options | global report automation
// Common Searches: how to embed a timeline image in PDF header using Aspose.Cells C# | Aspose.Cells render timeline as PNG for header | set centered header picture with Aspose.Cells PDF export | C# example adding PivotTable timeline to PDF header | align header image and title in Aspose.Cells PDF
// Developer Intent: Add a rendered PivotTable Timeline image to a PDF header and align it with a textual title.
// Use Cases: Quarterly sales PDF where the timeline filter appears beside the report title for quick period reference. | Project status report PDF that shows a Gantt‑style timeline graphic in the header next to the document name. | Invoice PDF that includes a payment‑date timeline image in the header for easy tracking of billing periods.
// AI Prompts: Generate code to render the timeline as a JPEG instead of PNG before setting it as the header picture. | Show how to place different timeline images in the left and right header sections of a PDF using Aspose.Cells. | Explain how to control the size, scaling, and vertical alignment of a header picture that contains a rendered timeline.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotFieldType

namespace AsposeCellsTimelineHeaderPdf
{
    // This example creates a workbook, adds sample data, builds a PivotTable, attaches a Timeline control, renders the worksheet (including the Timeline) to an in‑memory image, sets the image as the centered header picture while placing a text title on the left, and saves the result as a PDF where the header shows the timeline graphic aligned with the document title.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Populate sample data for a PivotTable (required for Timeline)
                worksheet.Cells["A1"].PutValue("Date");
                worksheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                worksheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
                worksheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));

                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["B2"].PutValue(1200);
                worksheet.Cells["B3"].PutValue(1500);
                worksheet.Cells["B4"].PutValue(1800);

                // 3. Add a PivotTable based on the data
                int pivotIdx = worksheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
                PivotTable pivot = worksheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivot.RefreshData();
                pivot.CalculateData();

                // 4. Add a Timeline linked to the PivotTable (using the "Date" field)
                worksheet.Timelines.Add(pivot, 10, 5, "Date");

                // 5. Render the worksheet (which now contains the Timeline) to an image in memory
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; explicit setting omitted to avoid missing property issue
                    OnePagePerSheet = true
                };
                SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
                using (MemoryStream imgStream = new MemoryStream())
                {
                    sheetRender.ToImage(0, imgStream);
                    byte[] headerImageBytes = imgStream.ToArray();

                    // 6. Set the rendered image as the center header picture
                    // Section index: 0=Left, 1=Center, 2=Right
                    worksheet.PageSetup.SetHeaderPicture(1, headerImageBytes);
                    // Use the image script "&G" to display the picture
                    worksheet.PageSetup.SetHeader(1, "&G");

                    // 7. Add a textual title to the left side of the header
                    worksheet.PageSetup.SetHeader(0, "Quarterly Sales Report");

                    // 8. Save the workbook as PDF with the header image visible
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        DisplayDocTitle = true // optional: show document title in PDF viewer
                    };
                    string outputPath = "QuarterlySalesReport.pdf";

                    // Ensure the output directory exists
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputPath, pdfOptions);
                    Console.WriteLine($"PDF generated at '{Path.GetFullPath(outputPath)}' with Timeline image in the header.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
