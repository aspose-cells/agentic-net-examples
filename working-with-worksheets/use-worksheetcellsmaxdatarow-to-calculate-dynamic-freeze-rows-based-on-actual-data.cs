// Title: Dynamically freeze rows up to the last populated row with Worksheet.Cells.MaxDataRow in Aspose.Cells for .NET
// AI Prompts: Write C# code that retrieves Worksheet.Cells.MaxDataRow and uses FreezePanes to lock all rows above that row in an Aspose.Cells workbook. | Show an example of calculating the maximum data row in a worksheet and applying a dynamic freeze pane based on that value with Aspose.Cells for .NET. | Extend the sample to also determine the maximum data column and freeze both rows and columns up to those limits using MaxDataColumn and FreezePanes.
// Common Searches: Aspose.Cells C# freeze panes based on the last data row | How to use Cells.MaxDataRow to set dynamic freeze rows in .NET | Freeze top rows automatically up to populated data with Aspose.Cells | C# code example for dynamic row freeze using MaxDataRow in Aspose.Cells
// Tags: dynamic row freeze Aspose.Cells | Worksheet.Cells.MaxDataRow example | FreezePanes based on data range .NET | calculate last populated row Aspose.Cells | C# Aspose.Cells freeze panes programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds sample data, obtains the index of the last row containing data via Worksheet.Cells.MaxDataRow, freezes all rows up to that index with FreezePanes, and saves the workbook as DynamicFreezeRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create workbook
            Worksheet sheet = workbook.Worksheets[0]; // get first worksheet

            // Example: populate some data (optional, replace with real data as needed)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(30);

            // Calculate the last row that contains data (zero‑based index)
            int maxDataRow = sheet.Cells.MaxDataRow;

            // Freeze rows up to the last data row.
            // FreezePanes(row, column, totalRows, totalColumns)
            // row/column specify the split location; totalRows/totalColumns specify how many rows/columns to freeze.
            sheet.FreezePanes(maxDataRow + 1, 0, maxDataRow + 1, 0);

            // Save the workbook with the dynamic freeze applied
            string outputPath = "DynamicFreezeRows.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
