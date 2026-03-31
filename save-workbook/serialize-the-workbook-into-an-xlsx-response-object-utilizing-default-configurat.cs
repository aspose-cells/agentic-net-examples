using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

public class Program
{
    public static void Main()
    {
        // 1. Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // 2. Populate the workbook with sample data (preserve original formatting)
        var sheet = workbook.Worksheets[0];
        sheet.Name = "SampleSheet";
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Data 1");
        sheet.Cells["A3"].PutValue("Data 2");

        // Example of applying a simple style to preserve formatting
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        sheet.Cells["A1"].SetStyle(style);

        // 3. Save the workbook to a file
        string filePath = "WorkbookExport.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to: {Path.GetFullPath(filePath)}");
    }
}