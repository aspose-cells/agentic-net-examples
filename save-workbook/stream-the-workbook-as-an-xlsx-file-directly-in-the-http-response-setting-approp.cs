using System;
using System.IO;
using Aspose.Cells;

public class WorkbookExportHandler
{
    public void ExportWorkbook(Stream outputStream)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

        workbook.Save(outputStream, SaveFormat.Xlsx);
    }
}

class Program
{
    static void Main()
    {
        var handler = new WorkbookExportHandler();
        using (var fileStream = new FileStream("SampleWorkbook.xlsx", FileMode.Create, FileAccess.Write))
        {
            handler.ExportWorkbook(fileStream);
        }

        Console.WriteLine("Workbook exported successfully.");
    }
}