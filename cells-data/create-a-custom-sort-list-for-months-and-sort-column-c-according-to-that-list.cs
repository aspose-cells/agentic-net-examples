// Title: Custom Month Sort in Excel with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to define a custom month order list and use Aspose.Cells' DataSorter to sort column C of an Excel worksheet while preserving headers, then saves the result as SortedByMonth.xlsx.
// Keywords: Aspose.Cells | C# | .NET | custom sort list | month sorting | DataSorter | Excel custom order | sort column by list | Excel automation | sample code
// Common Searches: Aspose.Cells sort column by custom month list C# | DataSorter AddKey custom list example | How to sort Excel months with Aspose.Cells .NET | Custom order sorting in Excel using Aspose | C# code for month based sorting in workbook
// Developer Intent: Apply a predefined month sequence to sort an Excel column using Aspose.Cells.
// Use Cases: Organize sales data by calendar month when months are stored as text. | Prepare monthly reports from unordered records without converting to dates. | Ensure month columns follow chronological order before creating pivot tables or charts.
// AI Prompts: Generate C# code that sorts an Excel worksheet column using a custom month list with Aspose.Cells. | Explain the parameters of DataSorter.AddKey for custom list sorting in Aspose.Cells. | Show how to sort multiple columns while keeping header rows intact using Aspose.Cells DataSorter.

using System;
using Aspose.Cells;

namespace CustomMonthSortExample
{
    // Demonstrates how to define a custom month order list and use Aspose.Cells' DataSorter to sort column C of an Excel worksheet while preserving headers, then saves the result as SortedByMonth.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Value");
            cells["C1"].PutValue("Month");

            // Sample data with months in random order
            cells["A2"].PutValue(1);
            cells["B2"].PutValue(100);
            cells["C2"].PutValue("March");

            cells["A3"].PutValue(2);
            cells["B3"].PutValue(200);
            cells["C3"].PutValue("January");

            cells["A4"].PutValue(3);
            cells["B4"].PutValue(150);
            cells["C4"].PutValue("December");

            cells["A5"].PutValue(4);
            cells["B5"].PutValue(120);
            cells["C5"].PutValue("July");

            cells["A6"].PutValue(5);
            cells["B6"].PutValue(180);
            cells["C6"].PutValue("May");

            // Define the custom month order list
            string monthOrder = "January,February,March,April,May,June,July,August,September,October,November,December";

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = true; // First row contains headers
            // Add custom sort key for column C (index 2) using the month order list
            sorter.AddKey(2, SortOrder.Ascending, monthOrder);

            // Define the range to sort (including headers)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 5,   // rows 0‑5 (A1:C6)
                EndColumn = 2 // columns A‑C
            };

            // Perform the sort
            sorter.Sort(worksheet.Cells, sortArea);

            // Save the workbook
            workbook.Save("SortedByMonth.xlsx");
        }
    }
}
