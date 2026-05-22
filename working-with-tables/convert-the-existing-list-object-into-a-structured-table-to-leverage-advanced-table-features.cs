using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains raw data
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the range that holds the data (including header row)
        // Adjust the range as needed for your source data
        string startCell = "A1";
        string endCell = "C5";

        // Convert the range into a structured table (ListObject)
        int tableIndex = worksheet.ListObjects.Add(startCell, endCell, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Configure table properties to leverage advanced features
        table.DisplayName = "MyStructuredTable";
        table.TableStyleType = TableStyleType.TableStyleMedium9; // Apply a built‑in style
        table.ShowTotals = true;                                   // Enable totals row
        table.ListColumns[0].TotalsCalculation = TotalsCalculation.Sum; // Example total

        // Save the workbook with the new table
        workbook.Save("output.xlsx");
    }
}