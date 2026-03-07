using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class TextCrossTypeDemo
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a long text into A1 that will exceed the cell width
        worksheet.Cells["A1"].PutValue("This is a very long text that will exceed the cell width and demonstrate TextCrossType behavior.");
        // Ensure the text is not shrunk to fit the cell
        worksheet.Cells["A1"].GetStyle().ShrinkToFit = false;

        // Make the column narrow so the text will cross cell boundaries
        worksheet.Cells.SetColumnWidth(0, 5);

        // Save images with different TextCrossType settings
        SaveImageWithTextCrossType(workbook, TextCrossType.Default, "TextCross_Default.png");
        SaveImageWithTextCrossType(workbook, TextCrossType.CrossKeep, "TextCross_CrossKeep.png");
        SaveImageWithTextCrossType(workbook, TextCrossType.CrossOverride, "TextCross_CrossOverride.png");
        SaveImageWithTextCrossType(workbook, TextCrossType.StrictInCell, "TextCross_StrictInCell.png");

        // Save a PDF using PdfSaveOptions with a specific TextCrossType
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.TextCrossType = TextCrossType.CrossKeep; // display crossed text and keep underlying cell text
        workbook.Save("TextCross_Pdf.pdf", pdfOptions);
    }

    // Helper method to render the first worksheet to an image using the specified TextCrossType
    static void SaveImageWithTextCrossType(Workbook workbook, TextCrossType type, string fileName)
    {
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = Aspose.Cells.Drawing.ImageType.Png;
        options.TextCrossType = type; // apply the desired TextCrossType

        // Render the first worksheet (index 0) to an image file
        SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], options);
        sheetRender.ToImage(0, fileName);
    }
}