// Title: C# – Aspose.Cells AutoFilter to Show Only Current‑Quarter Dates
// Description: Creates a workbook, populates column A with dates, defines an AutoFilter range, applies the dynamic filter ThisQuarter to the first column, refreshes the view to hide rows outside the current quarter, and saves the file as FilteredCurrentQuarter.xlsx.
// Keywords: Aspose.Cells C# AutoFilter | dynamic filter ThisQuarter | filter rows by current quarter | Excel date filter .NET | hide rows outside quarter Aspose | auto‑filter date column | Aspose.Cells filter example
// Common Searches: Aspose.Cells filter rows by current quarter C# | How to use DynamicFilterType.ThisQuarter in Aspose.Cells | C# code to hide dates not in this quarter with Aspose | AutoFilter date range current quarter Aspose.Cells .NET | Show only this quarter data in Excel using Aspose
// Developer Intent: Filter a worksheet so that only rows whose date values fall within the current calendar quarter remain visible.
// Use Cases: Quarterly sales or financial reports that automatically display only the current quarter’s entries. | Dashboard workbooks that refresh to show up‑to‑date quarter data without manual filtering. | Template files that hide outdated rows each time they are opened, ensuring users see only relevant quarter information.
// AI Prompts: Generate C# code with Aspose.Cells to filter rows for the previous quarter. | Show how to apply a custom date‑range AutoFilter in Aspose.Cells and export the filtered workbook. | Explain combining multiple DynamicFilterType values (e.g., ThisQuarter and ThisYear) on different columns in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, populates column A with dates, defines an AutoFilter range, applies the dynamic filter ThisQuarter to the first column, refreshes the view to hide rows outside the current quarter, and saves the file as FilteredCurrentQuarter.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data: a header and several dates
        worksheet.Cells["A1"].PutValue("Date");
        worksheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15)); // Q1
        worksheet.Cells["A3"].PutValue(new DateTime(2023, 4, 10)); // Q2
        worksheet.Cells["A4"].PutValue(DateTime.Now);               // Current date (in current quarter)
        worksheet.Cells["A5"].PutValue(DateTime.Now.AddMonths(1)); // Likely still in current quarter
        worksheet.Cells["A6"].PutValue(DateTime.Now.AddMonths(4)); // Outside current quarter

        // Define the auto‑filter range (header + data rows)
        worksheet.AutoFilter.Range = "A1:A6";

        // Apply a dynamic filter to keep only rows whose dates are in the current quarter
        worksheet.AutoFilter.DynamicFilter(0, DynamicFilterType.ThisQuarter);

        // Refresh the filter to hide rows that do not meet the criteria
        worksheet.AutoFilter.Refresh();

        // Save the filtered workbook
        workbook.Save("FilteredCurrentQuarter.xlsx");
    }
}
