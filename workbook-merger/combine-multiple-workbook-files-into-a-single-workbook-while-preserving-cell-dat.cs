using System;
using Aspose.Cells;

public class CombineWorkbooksDemo
{
    public static void Run()
    {
        string[] sourceFiles = new string[]
        {
            "File1.xlsx",
            "File2.xlsx",
            "File3.xlsx"
        };

        Workbook destination = new Workbook();
        destination.Worksheets.Clear();

        foreach (string filePath in sourceFiles)
        {
            Workbook source = new Workbook(filePath);
            destination.Combine(source);
        }

        destination.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        CombineWorkbooksDemo.Run();
    }
}