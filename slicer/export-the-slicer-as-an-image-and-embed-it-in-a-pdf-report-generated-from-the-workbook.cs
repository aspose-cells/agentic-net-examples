using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;      // PivotTable and related enums
using Aspose.Cells.Slicers;    // Slicer class

namespace AsposeCellsSample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate data for a pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Amount");
                dataSheet.Cells["A2"].PutValue("Fruit");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["A3"].PutValue("Vegetable");
                dataSheet.Cells["B3"].PutValue(80);
                dataSheet.Cells["A4"].PutValue("Fruit");
                dataSheet.Cells["B4"].PutValue(150);
                dataSheet.Cells["A5"].PutValue("Vegetable");
                dataSheet.Cells["B5"].PutValue(70);

                // Create a pivot table based on the data range A1:B5
                int pivotIdx = dataSheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivot = dataSheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field

                // Add a slicer linked to the pivot table (field index 0 = Category)
                int slicerIdx = dataSheet.Slicers.Add(pivot, 20, 2, 0);
                Slicer slicer = dataSheet.Slicers[slicerIdx];

                // Render the worksheet that contains the slicer to an image (PNG)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true // Render the whole sheet as a single page
                };

                using (MemoryStream slicerImageStream = new MemoryStream())
                {
                    // Render the sheet to the memory stream
                    SheetRender sheetRender = new SheetRender(dataSheet, imgOptions);
                    sheetRender.ToImage(0, slicerImageStream);
                    slicerImageStream.Position = 0; // Reset for reading

                    // Create a new worksheet to hold the slicer image
                    Worksheet imageSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                    imageSheet.Name = "SlicerImage";

                    // Insert the rendered image as a picture shape
                    imageSheet.Pictures.Add(0, 0, slicerImageStream);

                    // Save the workbook (including the picture) as a PDF report
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        ExportDocumentStructure = true
                    };
                    workbook.Save("SlicerReport.pdf", pdfOptions);
                }

                Console.WriteLine("PDF report generated successfully: SlicerReport.pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}