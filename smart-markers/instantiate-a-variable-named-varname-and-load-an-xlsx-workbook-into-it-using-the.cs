using System;
using Aspose.Cells;

public class WorkbookLoader
{
    public void Load()
    {
        // Path to the XLSX file to be loaded
        string filePath = "example.xlsx";

        // Load the workbook
        var workbook = new Workbook(filePath);
        Console.WriteLine($"Workbook loaded. Worksheets count: {workbook.Worksheets.Count}");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        new WorkbookLoader().Load();
    }
}