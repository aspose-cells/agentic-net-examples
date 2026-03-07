using System;
using System.IO;
using Aspose.Cells;

class ConvertTimelineTemplate
{
    static void Main()
    {
        string sourcePath = "template.xltm";
        string destPath = "output.xlsx";

        if (!File.Exists(sourcePath))
        {
            var wb = new Workbook();
            wb.Worksheets[0].Name = "Sheet1";
            wb.Worksheets[0].Cells["A1"].PutValue("Demo");
            wb.Save(sourcePath, SaveFormat.Xltm);
        }

        var workbook = new Workbook(sourcePath);
        workbook.Save(destPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
    }
}