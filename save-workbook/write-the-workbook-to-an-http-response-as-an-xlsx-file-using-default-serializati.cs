using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

public class WorkbookExporter
{
    public async Task ExportWorkbookAsync(Stream outputStream)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
        workbook.Save(outputStream, saveOptions);
        await outputStream.FlushAsync();
    }
}

public class Program
{
    public static async Task Main(string[] args)
    {
        var exporter = new WorkbookExporter();

        using var memoryStream = new MemoryStream();
        await exporter.ExportWorkbookAsync(memoryStream);

        // Reset stream position before saving to file
        memoryStream.Position = 0;
        await File.WriteAllBytesAsync("ExportedWorkbook.xlsx", memoryStream.ToArray());

        Console.WriteLine("Workbook exported successfully to ExportedWorkbook.xlsx");
    }
}