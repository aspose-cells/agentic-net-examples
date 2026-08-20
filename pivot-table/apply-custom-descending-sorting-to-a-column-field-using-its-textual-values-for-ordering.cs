// Title: Custom Descending Text Sort for a Worksheet Column with Aspose.Cells .NET
// Description: Demonstrates how to use Aspose.Cells' DataSorter to sort a worksheet column in descending order based on a predefined list of string values (e.g., Low > Medium > High). The example creates a workbook, adds a header and priority values, defines the custom order array, configures the sorter with headers, applies the sort to the range A1:A5, and saves the result as CustomSorted.xlsx.
// Keywords: Aspose.Cells custom sort | DataSorter descending text order | C# Excel custom list sort | sort column by predefined strings | pivot table custom sort order .NET
// Common Searches: Aspose.Cells sort column by custom string list | C# DataSorter custom descending order | How to apply custom text sorting in Excel with Aspose | Custom sort order for pivot table field .NET | Sort Excel range using predefined string hierarchy
// Developer Intent: Implement a descending sort on a worksheet column using a user‑defined list of textual values.
// Use Cases: Rank priority levels (Low, Medium, High) before generating reports. | Prepare data for a pivot table where business‑specific text ranking is required. | Export Excel files with status strings ordered according to a custom hierarchy.
// AI Prompts: Show C# code that uses Aspose.Cells DataSorter to sort a column by a custom descending list of strings. | Explain how to configure a custom text sort order for a pivot table field with Aspose.Cells .NET. | Provide step‑by‑step instructions to apply a predefined string hierarchy to an Excel range using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells' DataSorter to sort a worksheet column in descending order based on a predefined list of string values (e.g., Low > Medium > High). The example creates a workbook, adds a header and priority values, defines the custom order array, configures the sorter with headers, applies the sort to the range A1:A5, and saves the result as CustomSorted.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (including a header)
        worksheet.Cells["A1"].PutValue("Priority");
        worksheet.Cells["A2"].PutValue("Medium");
        worksheet.Cells["A3"].PutValue("Low");
        worksheet.Cells["A4"].PutValue("High");
        worksheet.Cells["A5"].PutValue("Medium");

        // Define a custom descending order for the textual values
        // (e.g., Low > Medium > High)
        string[] customDescendingOrder = new string[] { "Low", "Medium", "High" };

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true; // First row is a header
        // Add sorting key: column index 0 (A), descending, with the custom list
        sorter.AddKey(0, SortOrder.Descending, customDescendingOrder);

        // Perform the sort on the defined range
        sorter.Sort(worksheet.Cells, CellArea.CreateCellArea("A1", "A5"));

        // Save the result
        workbook.Save("CustomSorted.xlsx");
    }
}
