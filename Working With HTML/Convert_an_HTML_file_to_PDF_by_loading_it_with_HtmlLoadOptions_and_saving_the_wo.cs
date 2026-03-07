using System;
using Aspose.Cells;

class HtmlToPdfConverter
{
    static void Main()
    {
        string htmlFilePath = "input.html";
        string pdfFilePath = "output.pdf";

        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.SupportDivTag = true;

        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        workbook.Save(pdfFilePath, saveOptions);

        Console.WriteLine("Conversion completed: HTML -> PDF");
    }
}