using System;
using System.IO;
using Aspose.Cells;

public class WorkbookPdfResponse
{
    public void WriteWorkbookAsPdf(Stream outputStream)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Aspose.Cells PDF!");

        var pdfOptions = new PdfSaveOptions
        {
            EmbedStandardWindowsFonts = true
        };

        workbook.Save(outputStream, pdfOptions);
        outputStream.Flush();
    }

    public static void Main()
    {
        var response = new WorkbookPdfResponse();
        using (var fileStream = new FileStream("output.pdf", FileMode.Create, FileAccess.Write))
        {
            response.WriteWorkbookAsPdf(fileStream);
        }
        Console.WriteLine("PDF file 'output.pdf' has been created.");
    }
}