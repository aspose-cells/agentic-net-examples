using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using Aspose.Cells.Pivot;   // PivotTable and PivotFieldType are in this namespace

namespace AsposeCellsPivotToPdfImage
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a workbook and populate source data
                Workbook srcWb = new Workbook();
                Worksheet srcWs = srcWb.Worksheets[0];
                srcWs.Name = "Data";

                srcWs.Cells["A1"].PutValue("Category");
                srcWs.Cells["B1"].PutValue("Amount");
                srcWs.Cells["A2"].PutValue("Food");
                srcWs.Cells["B2"].PutValue(1200);
                srcWs.Cells["A3"].PutValue("Beverage");
                srcWs.Cells["B3"].PutValue(800);
                srcWs.Cells["A4"].PutValue("Snacks");
                srcWs.Cells["B4"].PutValue(450);

                // 2. Add a pivot table on a new worksheet
                Worksheet pivWs = srcWb.Worksheets.Add("Pivot");
                int pivIndex = pivWs.PivotTables.Add("=Data!A1:B4", "A3", "SalesPivot");
                PivotTable pivot = pivWs.PivotTables[pivIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // 3. Render the pivot worksheet to an image (PNG)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true
                };
                SheetRender sheetRender = new SheetRender(pivWs, imgOptions);
                string imagePath = "pivot.png";
                sheetRender.ToImage(0, imagePath); // render first page to file

                // 4. Create a new workbook to hold the static image
                Workbook pdfWb = new Workbook();
                Worksheet pdfWs = pdfWb.Worksheets[0];
                pdfWs.Name = "PivotImage";

                // 5. Insert the rendered image into the worksheet
                if (File.Exists(imagePath))
                {
                    try
                    {
                        using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                        {
                            // Add picture at cell A1 (row 0, column 0)
                            pdfWs.Pictures.Add(0, 0, imgStream);
                        }
                    }
                    catch (Exception exImg)
                    {
                        Console.WriteLine($"Image insertion error: {exImg.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("Rendered image file not found.");
                }

                // 6. Save the workbook as PDF – the pivot appears as a static image
                pdfWb.Save("PivotStaticImage.pdf");

                // Clean up temporary image file
                if (File.Exists(imagePath))
                {
                    try
                    {
                        File.Delete(imagePath);
                    }
                    catch (Exception exDel)
                    {
                        Console.WriteLine($"Failed to delete temporary image: {exDel.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}