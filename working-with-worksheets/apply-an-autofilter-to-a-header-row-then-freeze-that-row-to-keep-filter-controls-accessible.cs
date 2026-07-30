// Title: C# – Apply AutoFilter to Header Row and Freeze It with Aspose.Cells
// Description: Creates a workbook, writes a header and sample data, sets an AutoFilter on A1:C1, filters the Category column for "Fruit", refreshes the filter, freezes the first row using FreezePanes, and saves the result as an XLSX file.
// Keywords: Aspose.Cells AutoFilter C# | FreezePanes header row | filter and freeze worksheet | Aspose.Cells .NET example | programmatic Excel filter | freeze first row Aspose
// Common Searches: Aspose.Cells how to add AutoFilter to first row | C# freeze header after applying filter Aspose.Cells | AutoFilter and FreezePanes example .NET | keep filter dropdowns visible while scrolling Excel using Aspose | apply AutoFilter then freeze rows Aspose.Cells
// Developer Intent: Add an AutoFilter to a worksheet’s header row and keep that row fixed so filter controls stay visible during scrolling.
// Use Cases: Interactive product catalog where users can filter by category and the header remains in view. | Large financial reports that require frozen column titles and filter dropdowns for easy navigation. | Automated Excel dashboards that programmatically apply filters and lock the header for consistent user experience.
// AI Prompts: Generate C# code with Aspose.Cells to set an AutoFilter on A1:C1 and freeze the first row. | Show how to filter the "Category" column for a specific value and then lock the header using FreezePanes. | Explain step‑by‑step how AutoFilter.Refresh and FreezePanes work together to keep filter controls accessible in a large worksheet.

using System;
using Aspose.Cells;

namespace AutoFilterAndFreezeDemo
{
    // Creates a workbook, writes a header and sample data, sets an AutoFilter on A1:C1, filters the Category column for "Fruit", refreshes the filter, freezes the first row using FreezePanes, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header row (row 0) and some sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue("Fruit");
            cells["C2"].PutValue(1.2);

            cells["A3"].PutValue("Carrot");
            cells["B3"].PutValue("Vegetable");
            cells["C3"].PutValue(0.8);

            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue("Fruit");
            cells["C4"].PutValue(1.0);

            // Apply AutoFilter to the header row (A1:C1)
            worksheet.AutoFilter.Range = "A1:C1";

            // Example filter: show only rows where Category = "Fruit"
            worksheet.AutoFilter.Filter(1, "Fruit");
            worksheet.AutoFilter.Refresh();

            // Freeze the header row so filter controls stay visible while scrolling
            // Freeze at the second row (index 1) with 1 frozen row and 0 frozen columns
            worksheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("AutoFilterAndFreezeDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
