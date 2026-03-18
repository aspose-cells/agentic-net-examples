using System;
using System.IO;
using Aspose.Cells;

public class WorkbookExport
{
    public void ExportToStream(Stream outputStream)
    {
        // Create a new workbook and add sample data.
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Save workbook to the provided stream in XLSX format.
        workbook.Save(outputStream, SaveFormat.Xlsx);
    }
}

public class Program
{
    public static void Main()
    {
        using var memoryStream = new MemoryStream();
        var exporter = new WorkbookExport();
        exporter.ExportToStream(memoryStream);

        // Write the generated workbook to a file (optional)
        File.WriteAllBytes("output.xlsx", memoryStream.ToArray());

        Console.WriteLine("Workbook exported successfully.");
    }
}