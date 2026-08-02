// Title: Aspose.Cells .NET Pivot Table – Display Null Values as "N/A"
// Description: Learn how to create a workbook with sample sales data, add a pivot table, enable custom null handling, and set the NullString property to "N/A" using Aspose.Cells for C#. The example refreshes the pivot, saves the file, and demonstrates a clean way to show missing values in Excel reports.
// Keywords: Aspose.Cells pivot table null string | Display null as N/A Aspose.Cells .NET | C# Aspose.Cells custom null display | PivotTable NullString property | Excel pivot missing data placeholder | Aspose.Cells example C# | Replace empty cells with text in pivot
// Common Searches: Aspose.Cells set NullString for pivot table | C# show N/A for null values in Excel pivot | How to display missing data as N/A in Aspose.Cells | Pivot table custom null text .NET | Aspose.Cells replace empty cells with placeholder
// Developer Intent: The developer needs a .NET solution that makes null or empty cells appear as the text "N/A" in an Aspose.Cells‑generated Excel pivot table.
// Use Cases: Create sales dashboards where blank product names or sales figures are clearly marked as "N/A". | Automate client‑ready Excel reports that replace database nulls with a readable placeholder in pivot analyses. | Export data from ERP systems to Excel and ensure any missing values are visible as "N/A" in the pivot view.
// AI Prompts: Generate C# code that builds a pivot table with Aspose.Cells, enables DisplayNullString, sets NullString to "N/A", refreshes the pivot, and saves the workbook. | Explain step‑by‑step how to configure Aspose.Cells pivot tables to replace null values with a custom string and why this improves report readability. | Provide best practices for handling null source data in Excel pivots using Aspose.Cells, including performance tips and localization considerations.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsNullDisplayDemo
{
    // Learn how to create a workbook with sample sales data, add a pivot table, enable custom null handling, and set the NullString property to "N/A" using Aspose.Cells for C#. The example refreshes the pivot, saves the file, and demonstrates a clean way to show missing values in Excel reports.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with null values
            // Header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");

            // Data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(null); // Null sales value

            sheet.Cells["A4"].PutValue(null); // Null product name
            sheet.Cells["B4"].PutValue(800);

            // Add a pivot table based on the data range A1:B4, place it at C3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Sales

            // Enable custom display for null values and set the desired string
            pivotTable.DisplayNullString = true;
            pivotTable.NullString = "N/A";

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_NullDisplay.xlsx");

            // Optional: Inform the user
            Console.WriteLine("Workbook saved as 'PivotTable_NullDisplay.xlsx' with null values displayed as 'N/A'.");
        }
    }
}
