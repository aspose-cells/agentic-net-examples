// Title: Copy only visible rows from an AutoFiltered range to another worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells CopyRows together with PasteOptions.OnlyVisibleCells to duplicate only the rows that remain visible after applying an AutoFilter in C#. | Write C# code that applies an AutoFilter on a column, then copies the rows that stay displayed to a new sheet while skipping hidden rows. | Generate a .NET snippet that builds a workbook, sets an AutoFilter, and transfers only the rows that are shown after filtering to another worksheet using Aspose.Cells.
// Common Searches: asp.net copy rows that remain visible after autofilter to a new worksheet using Aspose.Cells | c# copy only visible rows after applying AutoFilter with PasteOptions in Aspose.Cells | how to duplicate filtered rows to another sheet in Aspose.Cells for .NET | example of copying filtered data to a separate worksheet using Aspose.Cells
// Tags: CopyRows filtered rows Aspose.Cells | AutoFilter row copy C# | copy filtered rows to new worksheet .NET | PasteOptions visible cells Aspose.Cells | Aspose.Cells copy visible rows example

using System;
using Aspose.Cells;

// The example creates a workbook, fills a small table, applies an AutoFilter to show only rows where the Category equals "Apple", and then copies all rows from the source sheet to a newly added sheet using CopyRows with PasteOptions.OnlyVisibleCells set to true. This ensures that only the visible (filtered) rows are duplicated. The resulting workbook is saved as CopyVisibleRowsDemo.xlsx.
class CopyVisibleRowsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        Cells sourceCells = sourceSheet.Cells;

        // Populate header and sample data
        sourceCells["A1"].PutValue("Category");
        sourceCells["B1"].PutValue("Value");
        sourceCells["A2"].PutValue("Apple");
        sourceCells["B2"].PutValue(10);
        sourceCells["A3"].PutValue("Banana");
        sourceCells["B3"].PutValue(20);
        sourceCells["A4"].PutValue("Apple");
        sourceCells["B4"].PutValue(30);
        sourceCells["A5"].PutValue("Cherry");
        sourceCells["B5"].PutValue(40);
        sourceCells["A6"].PutValue("Apple");
        sourceCells["B6"].PutValue(50);

        // Apply AutoFilter on the header row covering columns A‑B
        sourceSheet.AutoFilter.Range = "A1:B6";

        // Filter to show only rows where Category = "Apple"
        sourceSheet.AutoFilter.AddFilter(0, "Apple");
        sourceSheet.AutoFilter.Refresh(); // hides non‑matching rows

        // Add a destination worksheet where rows will be copied
        Worksheet destSheet = workbook.Worksheets.Add("Copy");
        Cells destCells = destSheet.Cells;

        // Set up copy and paste options; paste only visible cells
        CopyOptions copyOptions = new CopyOptions(); // default options
        PasteOptions pasteOptions = new PasteOptions();
        pasteOptions.OnlyVisibleCells = true; // copy only visible rows

        // Copy all rows from source to destination using the options
        destCells.CopyRows(
            sourceCells,
            0,                                 // source start row (including header)
            0,                                 // destination start row
            sourceCells.MaxDisplayRange.RowCount, // number of rows to copy
            copyOptions,
            pasteOptions);

        // Verify that only visible rows were duplicated
        Console.WriteLine("Destination sheet values after copying visible rows:");
        for (int r = 0; r < destCells.MaxDisplayRange.RowCount; r++)
        {
            string category = destCells[r, 0].StringValue;
            string value = destCells[r, 1].StringValue;
            if (!string.IsNullOrEmpty(category) || !string.IsNullOrEmpty(value))
            {
                Console.WriteLine($"Row {r + 1}: Category = {category}, Value = {value}");
            }
        }

        // Save the workbook
        workbook.Save("CopyVisibleRowsDemo.xlsx");
    }
}
