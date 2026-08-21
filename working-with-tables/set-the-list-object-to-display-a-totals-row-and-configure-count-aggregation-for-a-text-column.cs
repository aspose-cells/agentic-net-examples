// Title: Add a Totals Row with Count Aggregation for a Text Column using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a ListObject (table), enables its totals row, sets the first text column to TotalsCalculation.Count, optionally adds a label, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# ListObject totals row | ListObject TotalsCalculation.Count | Excel table count aggregation | Aspose.Cells table totals | ShowTotals property | ListColumns TotalsRowLabel
// Common Searches: Aspose.Cells show totals row | C# ListObject count total | How to set TotalsCalculation.Count in Aspose.Cells | Add totals row to Excel table using Aspose.Cells | Label totals row column Aspose.Cells
// Developer Intent: Add a totals row to a ListObject and configure the first (text) column to display a Count aggregation.
// Use Cases: Generate a summary row that counts distinct category entries in an automated Excel report. | Create a financial sheet where the totals row shows a custom label and count for a non‑numeric column. | Build a data‑export routine that adds a totals row with count statistics for textual data before distribution.
// AI Prompts: Write C# code with Aspose.Cells to add a ListObject, enable its totals row, set TotalsCalculation.Count for the first column, and assign a custom label. | Explain how TotalsCalculation.Count works for a text column in an Aspose.Cells ListObject and how to read the calculated count after saving the workbook. | Provide a C# example that applies different TotalsCalculation types (Sum, Average, Count) to multiple ListObject columns using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, defines a ListObject (table), enables its totals row, sets the first text column to TotalsCalculation.Count, optionally adds a label, and saves the file as an Excel workbook.
class ListObjectTotalsCountDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data: a text column and a numeric column
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("A");
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["A5"].PutValue("B");
            worksheet.Cells["B5"].PutValue(40);

            // Add a ListObject (table) that includes the data range
            int tableIndex = worksheet.ListObjects.Add("A1", "B5", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Enable the totals row for the table
            table.ShowTotals = true;

            // Configure the totals row to display a Count aggregation for the text column (first column)
            table.ListColumns[0].TotalsCalculation = TotalsCalculation.Count;

            // Optionally set a label for the totals cell of the first column
            table.ListColumns[0].TotalsRowLabel = "Count";

            // Save the workbook to a file
            workbook.Save("ListObjectTotalsCountDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        ListObjectTotalsCountDemo.Run();
    }
}
