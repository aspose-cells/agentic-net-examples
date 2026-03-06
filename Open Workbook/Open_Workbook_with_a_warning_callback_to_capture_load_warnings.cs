using System;
using System.Collections.Generic;
using Aspose.Cells;

public class LoadWarningCollector : IWarningCallback
{
    public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

    public void Warning(WarningInfo warningInfo)
    {
        Warnings.Add(warningInfo);
        Console.WriteLine($"Warning captured: {warningInfo.Type} - {warningInfo.Description}");
    }
}

public class LoadWorkbookWithWarningsDemo
{
    public static void Run()
    {
        string filePath = "sample.xlsx";

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.WarningCallback = new LoadWarningCollector();

        Workbook workbook = new Workbook(filePath, loadOptions);

        Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");

        var collector = (LoadWarningCollector)loadOptions.WarningCallback;
        Console.WriteLine($"Total warnings captured during load: {collector.Warnings.Count}");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        LoadWorkbookWithWarningsDemo.Run();
    }
}