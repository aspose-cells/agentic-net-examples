// Title: Read the default Grand Total label from a PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: how to read grand total row label from a pivot table with Aspose.Cells in C# | Aspose.Cells C# retrieve default grand total text after pivot calculation | programmatically get pivot table total label using Aspose.Cells .NET | C# example to extract grand total label from an Excel pivot with Aspose.Cells
// Tags: Aspose.Cells read pivot grand total label | C# Aspose.Cells get pivot total row text | Aspose.Cells calculate pivot and extract label | Aspose.Cells create pivot from data range | Aspose.Cells save workbook as Excel file

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This C# example creates a workbook, populates sample data, adds a PivotTable, configures row and data fields, refreshes and calculates the pivot, determines the Grand Total row position, reads its default label from the worksheet, prints the label, and saves the file as PivotCustomNames.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and fill it with sample data.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("A");
            cells["B4"].PutValue(30);
            cells["A5"].PutValue("B");
            cells["B5"].PutValue(40);

            // Add a pivot table based on the data range.
            string sourceRange = "A1:B5";
            int pivotRow = 7, pivotColumn = 0;
            int pivotIndex = ws.PivotTables.Add(sourceRange, pivotRow, pivotColumn, "PivotTable1");
            PivotTable pt = ws.PivotTables[pivotIndex];

            // Configure the pivot: Category as row field, Amount as data field.
            pt.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pt.AddFieldToArea(PivotFieldType.Data, 1);  // Amount

            // Refresh and calculate the pivot table.
            pt.RefreshData();
            pt.CalculateData();

            // Output the grand total label (default label).
            // The grand total row appears after the row fields; calculate its position.
            int grandTotalRow = pivotRow + pt.RowFields.Count + 1;
            string grandTotalLabel = ws.Cells[grandTotalRow, pivotColumn].StringValue;
            Console.WriteLine("Grand Total label: " + grandTotalLabel);

            // Save the workbook.
            string outputPath = "PivotCustomNames.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
