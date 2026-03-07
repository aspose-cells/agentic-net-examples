using Aspose.Cells;
using System;

public class ClearAllPageBreaks
{
    public static void Run()
    {
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        worksheet.HorizontalPageBreaks.Add(5);
        worksheet.VerticalPageBreaks.Add(2);

        worksheet.HorizontalPageBreaks.Clear();

        while (worksheet.VerticalPageBreaks.Count > 0)
        {
            worksheet.VerticalPageBreaks.RemoveAt(0);
        }

        workbook.Save("ClearedPageBreaks.xlsx");
    }
}

public class Program
{
    public static void Main()
    {
        ClearAllPageBreaks.Run();
    }
}