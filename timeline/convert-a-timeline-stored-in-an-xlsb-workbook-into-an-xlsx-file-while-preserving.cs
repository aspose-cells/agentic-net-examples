using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

public class Program
{
    public static void Main()
    {
        TimelineConversion.Run();
    }
}

public class TimelineConversion
{
    public static void Run()
    {
        string sourceFile = "timeline.xlsb";
        string destinationFile = "timeline_converted.xlsx";

        ConversionUtility.Convert(sourceFile, destinationFile);

        Workbook workbook = new Workbook(destinationFile);
        Console.WriteLine($"Conversion completed successfully. Worksheets count: {workbook.Worksheets.Count}");
    }
}