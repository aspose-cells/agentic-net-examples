using System;
using System.IO;
using Aspose.Cells;

public class WorkbookPdfExporter
{
    public void Export(Stream outputStream)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        var pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        workbook.Save(outputStream, pdfOptions);
    }
}

public class Program
{
    public static void Main()
    {
        using (var fs = new FileStream("output.pdf", FileMode.Create, FileAccess.Write))
        {
            var exporter = new WorkbookPdfExporter();
            exporter.Export(fs);
        }
        Console.WriteLine("PDF exported successfully.");
    }
}