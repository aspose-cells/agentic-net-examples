// Title: Export Pivot Table as a Static Image in a PDF using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, build a pivot table, render the pivot sheet to a PNG image with SheetRender, embed the image in a new worksheet, and save the result as a PDF. The generated PDF contains a non‑editable snapshot of the pivot table.
// Keywords: Aspose.Cells | C# | .NET | pivot table | export to PDF | render to image | SheetRender | static snapshot | Excel to PDF | PDF image embedding
// Common Searches: Aspose.Cells export pivot table as image PDF C# | Render Excel pivot table to PNG and save as PDF | How to embed a pivot snapshot in a PDF using Aspose.Cells | Convert pivot sheet to image for PDF output .NET | Static pivot table image in PDF Aspose
// Developer Intent: Create a PDF that displays a fixed image of the current pivot table view.
// Use Cases: Distribute a read‑only view of a pivot report to stakeholders. | Archive pivot analysis for compliance without exposing source data. | Add a visual pivot snapshot to automated PDF dashboards or reports.
// AI Prompts: Show how to export several pivot tables as separate images within one PDF. | Modify the example to use JPEG instead of PNG and adjust scaling before embedding. | Add a caption or title below the pivot image in the generated PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsPivotToPdfImageDemo
{
    // Demonstrates how to create a workbook, build a pivot table, render the pivot sheet to a PNG image with SheetRender, embed the image in a new worksheet, and save the result as a PDF. The generated PDF contains a non‑editable snapshot of the pivot table.
    public class Program
    {
        public static void Main()
        {
            // Create a workbook and add sample data
            Workbook sourceWb = new Workbook();
            Worksheet dataSheet = sourceWb.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Drink");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Food");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Drink");
            dataSheet.Cells["B5"].PutValue(70);

            // Add a worksheet for the pivot table
            Worksheet pivotSheet = sourceWb.Worksheets.Add("Pivot");
            // Create the pivot table
            int pivotIdx = pivotSheet.PivotTables.Add("=Data!A1:B5", "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIdx];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            // Refresh to ensure the pivot reflects the latest data
            sourceWb.Worksheets.RefreshPivotTables();

            // Render the pivot sheet to an image (captures current state)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };
            SheetRender sheetRender = new SheetRender(pivotSheet, imgOptions);
            using (MemoryStream imgStream = new MemoryStream())
            {
                // Render first (and only) page of the sheet to the stream
                sheetRender.ToImage(0, imgStream);
                imgStream.Position = 0; // Reset stream position for reading

                // Create a new workbook that will hold the image
                Workbook pdfWb = new Workbook();
                Worksheet imgSheet = pdfWb.Worksheets[0];
                imgSheet.Name = "PivotImage";

                // Insert the rendered image into the worksheet as a picture
                // (0,0) specifies the upper‑left cell where the picture will be placed
                imgSheet.Pictures.Add(0, 0, imgStream);

                // Save the workbook as PDF – the picture (static image of the pivot) will be embedded
                pdfWb.Save("PivotTable_As_Image.pdf");
            }

            Console.WriteLine("Pivot table exported as static image inside PDF successfully.");
        }
    }
}
