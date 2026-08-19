// Title: Export a Pivot Table Slicer as PNG and embed it in a PDF report with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook with a pivot table and slicer, render the slicer to a PNG image, insert the image into a worksheet, and save the workbook as a PDF report using Aspose.Cells for .NET.
// Keywords: Aspose.Cells slicer export | C# render slicer to PNG | embed slicer image in PDF | pivot table slicer PDF report | Aspose.Cells PDFSaveOptions | SheetRender ToImage C# | Aspose.Cells tutorial
// Common Searches: export slicer as image Aspose.Cells .NET | how to embed slicer PNG in PDF using Aspose.Cells | render worksheet with slicer to PNG C# | save workbook as PDF with slicer image | Aspose.Cells slicer to PDF example
// Developer Intent: Create a PDF report that includes a static image of a pivot table slicer generated with Aspose.Cells.
// Use Cases: Generate printable PDFs where interactive slicers are represented as static images for distribution. | Automate reporting pipelines that capture slicer visuals for archival or compliance purposes. | Customize the layout of PDF reports by positioning slicer images on dedicated worksheets.
// AI Prompts: Provide C# code using Aspose.Cells to export a slicer to PNG and embed it in a PDF. | Explain how to render a worksheet containing a slicer to an image before converting the workbook to PDF. | Suggest ways to control the placement and size of the slicer image on the PDF page.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Saving;

// Demonstrates how to create a workbook with a pivot table and slicer, render the slicer to a PNG image, insert the image into a worksheet, and save the workbook as a PDF report using Aspose.Cells for .NET.
class ExportSlicerToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Fruit");
        sheet.Cells["A3"].PutValue("Vegetable");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a pivot table based on the data range
        int pivotIdx = sheet.PivotTables.Add("A1:B3", "D1", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

        // Add a slicer linked to the pivot table (field index 0 = Category)
        int slicerIdx = sheet.Slicers.Add(pivot, 20, 0, 0);
        Slicer slicer = sheet.Slicers[slicerIdx];
        slicer.IsPrintable = true; // make slicer printable (optional)

        // Render the worksheet (which now contains the slicer) to an image file
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageType = Aspose.Cells.Drawing.ImageType.Png;
        SheetRender sheetRender = new SheetRender(sheet, imgOptions);
        string slicerImagePath = "slicer.png";
        sheetRender.ToImage(0, slicerImagePath); // uses SheetRender.ToImage(int, string) rule

        // Insert the rendered slicer image into a new worksheet for PDF embedding
        Worksheet imageSheet = workbook.Worksheets.Add("SlicerImage");
        imageSheet.Pictures.Add(0, 0, slicerImagePath); // place image at top‑left corner

        // Save the workbook as a PDF report, embedding the slicer image
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true;
        workbook.Save("ReportWithSlicer.pdf", pdfOptions);
    }
}
