// Title: C# – Use Aspose.Cells AutoFilter to show only rows from the current quarter
// Description: Creates a workbook, adds a date column, defines an AutoFilter range, applies DynamicFilterType.ThisQuarter to keep rows whose dates fall in the current quarter, refreshes the view, and saves the file as FilteredCurrentQuarter.xlsx.
// Keywords: Aspose.Cells | C# | .NET | AutoFilter | DynamicFilterType.ThisQuarter | filter current quarter | Excel date filter | hide rows by quarter | quarterly report automation
// Common Searches: Aspose.Cells filter rows by current quarter C# | DynamicFilterType.ThisQuarter example | AutoFilter this quarter Aspose.Cells .NET | Show only current quarter dates in Excel using C# | Hide rows outside current quarter with Aspose.Cells
// Developer Intent: Apply an AutoFilter that leaves visible only the rows whose date values belong to the current calendar quarter.
// Use Cases: Generate a quarterly financial report that automatically excludes data from other quarters. | Build a dashboard that displays only the current quarter's transactions without manual date selection. | Prepare a workbook for a quarterly review by programmatically applying the ThisQuarter filter before distribution.
// AI Prompts: Write C# code with Aspose.Cells to filter rows for the previous quarter instead of the current one. | Extend the example to also filter a numeric column for values between 1000 and 5000 while keeping the current‑quarter date filter. | Explain how DynamicFilterType.ThisQuarter works internally and how it can be combined with custom date ranges in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, adds a date column, defines an AutoFilter range, applies DynamicFilterType.ThisQuarter to keep rows whose dates fall in the current quarter, refreshes the view, and saves the file as FilteredCurrentQuarter.xlsx.
class FilterCurrentQuarter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data: header + dates
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 20));
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 10));
        sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 5));
        sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 12));
        sheet.Cells["A7"].PutValue(DateTime.Now); // current date (in current quarter)

        // Define the auto‑filter range (including header and data rows)
        sheet.AutoFilter.Range = "A1:A7";

        // Apply a dynamic filter to keep only dates in the current quarter
        sheet.AutoFilter.DynamicFilter(0, DynamicFilterType.ThisQuarter);

        // Refresh the filter to hide rows that do not meet the criteria
        sheet.AutoFilter.Refresh();

        // Save the workbook
        workbook.Save("FilteredCurrentQuarter.xlsx");
    }
}
