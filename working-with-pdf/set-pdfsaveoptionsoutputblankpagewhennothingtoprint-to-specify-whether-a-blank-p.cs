using System;
using System.IO;
using Aspose.Cells;

class SetOutputBlankPageDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add an extra empty worksheet to keep at least one visible sheet
        Worksheet emptySheet = workbook.Worksheets.Add("EmptySheet");

        // Add some content to the visible sheet so that there is something to print
        emptySheet.Cells["A1"].PutValue("Sample Data");

        // Hide the first worksheet to simulate a scenario with nothing to print on that sheet
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.IsVisible = false;

        // Initialize PDF save options without outputting a blank page
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OutputBlankPageWhenNothingToPrint = false
        };

        // Save the workbook to PDF without a blank page
        string outputPathNoBlank = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "NoBlankPage.pdf");
        workbook.Save(outputPathNoBlank, pdfOptions);

        // Change the option to output a blank page when there is nothing to print
        pdfOptions.OutputBlankPageWhenNothingToPrint = true;

        // Save the workbook to PDF with a blank page
        string outputPathWithBlank = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "WithBlankPage.pdf");
        workbook.Save(outputPathWithBlank, pdfOptions);
    }
}