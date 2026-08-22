// Title: Clear the AutoFilter on column AJ and reveal all rows using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to remove the AutoFilter from column AJ and refresh the worksheet so all hidden rows become visible. | Programmatically delete the filter criteria applied to column index 35 with Aspose.Cells, then save the workbook without any active filters. | Write C# code that calls Worksheet.AutoFilter.RemoveFilter for column AJ, invokes Refresh, and exports the result as an .xlsx file.
// Common Searches: Aspose.Cells C# how to remove AutoFilter from a specific column and show hidden rows | clear filter on column AJ in Excel using Aspose.Cells .NET | refresh worksheet after deleting AutoFilter criteria with Aspose.Cells C# example | remove column filter programmatically Aspose.Cells and unhide rows | Aspose.Cells remove AutoFilter criteria column index 35
// Tags: Aspose.Cells remove column filter C# | remove AutoFilter from column AJ Aspose.Cells | unhide rows after clearing Excel filter Aspose | refresh worksheet AutoFilter Aspose.Cells | Excel column index 35 filter removal .NET

using System;
using Aspose.Cells;

namespace ClearAutoFilterOnColumnAJ
{
    // The C# program loads an Excel workbook, removes any AutoFilter applied to column AJ (index 35) using Worksheet.AutoFilter.RemoveFilter, refreshes the filter to unhide all rows, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Column AJ corresponds to index 35 (0‑based)
            int columnIndexAJ = 35;

            // Remove any filter applied to column AJ
            worksheet.AutoFilter.RemoveFilter(columnIndexAJ);

            // Refresh the autofilter to apply the change and unhide rows
            worksheet.AutoFilter.Refresh();

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
