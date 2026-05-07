using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class NonSequentialRangeDemo
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        for (int i = 0; i < 15; i++)
        {
            cells[i, 0].PutValue($"Row{i + 1}");
            cells[i, 1].PutValue(i * 10);
        }

        string rangeSpec = "A1:A3, C5:D6, E10";

        UnionRange unionRange = BuildUnionRange(cells, rangeSpec);

        Style highlight = workbook.CreateStyle();
        highlight.ForegroundColor = System.Drawing.Color.Yellow;
        highlight.Pattern = BackgroundType.Solid;
        unionRange.SetStyle(highlight);

        workbook.Save("NonSequentialRangeDemo.xlsx");
    }

    static UnionRange BuildUnionRange(Cells cells, string spec)
    {
        string[] parts = spec.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        List<AsposeRange> ranges = new List<AsposeRange>();

        foreach (string part in parts)
        {
            ranges.Add(cells.CreateRange(part.Trim()));
        }

        if (ranges.Count == 0)
            throw new ArgumentException("Range specification must contain at least one range.");

        AsposeRange baseRange = ranges[0];
        AsposeRange[] additional = ranges.Skip(1).ToArray();

        return baseRange.UnionRanges(additional);
    }
}