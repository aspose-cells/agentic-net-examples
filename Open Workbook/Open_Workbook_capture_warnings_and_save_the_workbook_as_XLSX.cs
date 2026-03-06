using System;
using System.Collections.Generic;
using Aspose.Cells;

public class WarningCollector : IWarningCallback
{
    public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

    public void Warning(WarningInfo warningInfo)
    {
        Warnings.Add(warningInfo);
        Console.WriteLine($"Warning: {warningInfo.Type} - {warningInfo.Description}");
    }
}

public class OpenCaptureSaveDemo
{
    public static void Run()
    {
        string sourcePath = "input.xlsx";

        var warningCallback = new WarningCollector();

        Workbook workbook = new Workbook(sourcePath);
        workbook.Settings.WarningCallback = warningCallback;

        Worksheet sheet = workbook.Worksheets[0];
        int idx1 = workbook.Worksheets.Names.Add("DemoRange");
        workbook.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1";

        int idx2 = workbook.Worksheets.Names.Add("DemoRange");
        workbook.Worksheets.Names[idx2].RefersTo = "=Sheet1!$A$1";

        workbook.Save("output.xlsx", SaveFormat.Xlsx);

        Console.WriteLine($"Total warnings captured: {warningCallback.Warnings.Count}");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        OpenCaptureSaveDemo.Run();
    }
}