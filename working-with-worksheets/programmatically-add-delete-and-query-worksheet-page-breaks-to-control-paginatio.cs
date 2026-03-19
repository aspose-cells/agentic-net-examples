using System;
using Aspose.Cells;

public class PageBreakDemo
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data (50 rows x 20 columns)
        for (int r = 0; r < 50; r++)
        {
            for (int c = 0; c < 20; c++)
            {
                sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
            }
        }

        // ---------- Add Horizontal Page Breaks ----------
        // Add a break after row index 10 (zero‑based)
        sheet.HorizontalPageBreaks.Add(10);
        // Add a break after row index 20
        sheet.HorizontalPageBreaks.Add(20);
        // Add a break using a cell address (D35 -> row 34)
        int rowD35, colD35;
        CellsHelper.CellNameToIndex("D35", out rowD35, out colD35);
        sheet.HorizontalPageBreaks.Add(rowD35);

        // ---------- Add Vertical Page Breaks ----------
        // Add a break after column index 5
        sheet.VerticalPageBreaks.Add(5);
        // Add a break after column index 0, starting at row 10 (default end row is the last row)
        sheet.VerticalPageBreaks.Add(0, 10);
        // Add a break using a cell address (G2 -> column 6)
        int rowG2, colG2;
        CellsHelper.CellNameToIndex("G2", out rowG2, out colG2);
        sheet.VerticalPageBreaks.Add(colG2);

        // ---------- Delete Specific Page Breaks ----------
        // Remove the second horizontal page break (index 1) if it exists
        if (sheet.HorizontalPageBreaks.Count > 1)
        {
            sheet.HorizontalPageBreaks.RemoveAt(1);
        }

        // Remove the first vertical page break (index 0) if it exists
        if (sheet.VerticalPageBreaks.Count > 0)
        {
            sheet.VerticalPageBreaks.RemoveAt(0);
        }

        // ---------- Query and Output Remaining Page Breaks ----------
        Console.WriteLine("Horizontal Page Breaks (row indices):");
        foreach (HorizontalPageBreak hpb in sheet.HorizontalPageBreaks)
        {
            Console.WriteLine($"Row: {hpb.Row}");
        }

        Console.WriteLine("Vertical Page Breaks (column indices):");
        foreach (VerticalPageBreak vpb in sheet.VerticalPageBreaks)
        {
            Console.WriteLine($"Column: {vpb.Column}, StartRow: {vpb.StartRow}, EndRow: {vpb.EndRow}");
        }

        // Save the workbook with the configured page breaks
        workbook.Save("PageBreakDemo.xlsx");
    }
}