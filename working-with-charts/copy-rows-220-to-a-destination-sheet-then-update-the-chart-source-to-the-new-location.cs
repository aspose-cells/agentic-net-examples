using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // ---------- Create source workbook with data and a chart ----------
        Workbook sourceWb = new Workbook();
        Worksheet sourceSheet = sourceWb.Worksheets[0];
        sourceSheet.Name = "Source";

        // Fill rows 2‑20 (indexes 1‑19) in columns A and B
        for (int i = 1; i <= 19; i++)
        {
            sourceSheet.Cells[i, 0].PutValue($"Item {i}");
            sourceSheet.Cells[i, 1].PutValue(i * 10);
        }

        // Add a chart that uses the source data range A2:B20
        int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 0, 2, 15, 7);
        Chart chart = sourceSheet.Charts[chartIdx];
        chart.NSeries.Add("=Source!$A$2:$B$20", true);
        chart.Title.Text = "Sample Chart";

        // ---------- Create destination workbook ----------
        Workbook destWb = new Workbook();
        Worksheet destSheet = destWb.Worksheets[0];
        destSheet.Name = "Destination";

        // ---------- Copy rows 2‑20 from source to destination ----------
        CopyOptions copyOptions = new CopyOptions();
        copyOptions.ReferToDestinationSheet = true; // ensures references point to destination when applicable

        // Copy 19 rows starting at source row index 1 to destination row index 1
        destSheet.Cells.CopyRows(sourceSheet.Cells, 1, 1, 19, copyOptions);

        // ---------- Update the chart to refer to the new range on the destination sheet ----------
        // Add a new series that points to Destination!A2:B20
        chart.NSeries.Add("=Destination!$A$2:$B$20", true);

        // ---------- Save workbooks ----------
        sourceWb.Save("SourceWorkbook.xlsx");
        destWb.Save("DestinationWorkbook.xlsx");
    }
}