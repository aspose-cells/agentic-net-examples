using System;
using System.IO;
using Aspose.Cells;

public class WorkbookSerializer
{
    public static byte[] SerializeToXlsx()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            return stream.ToArray();
        }
    }
}

public class Program
{
    public static void Main()
    {
        byte[] data = WorkbookSerializer.SerializeToXlsx();
        File.WriteAllBytes("output.xlsx", data);
        Console.WriteLine($"Workbook serialized, size: {data.Length} bytes");
    }
}