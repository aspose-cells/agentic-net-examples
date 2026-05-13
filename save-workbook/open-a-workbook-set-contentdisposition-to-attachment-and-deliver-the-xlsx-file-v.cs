using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

public class WorkbookExporter
{
    public void ExportWorkbook(Stream outputStream, string fileName = "output.xlsx")
    {
        var workbook = new Workbook("input.xlsx");
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
        workbook.Save(outputStream, saveOptions);
    }
}

public class Program
{
    public static void Main()
    {
        using (var fs = new FileStream("output.xlsx", FileMode.Create, FileAccess.Write))
        {
            var exporter = new WorkbookExporter();
            exporter.ExportWorkbook(fs);
        }
        Console.WriteLine("Export completed.");
    }
}