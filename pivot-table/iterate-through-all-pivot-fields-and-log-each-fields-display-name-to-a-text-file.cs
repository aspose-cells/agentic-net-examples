// Title: Log each pivot table field’s display name to a text file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates over a PivotTable's BaseFields collection and saves each field's DisplayName into a .txt file. | Create a reusable method that accepts a PivotTable object and a file path, then writes all pivot field names to the specified text file using Aspose.Cells. | Extend the logging routine to also record each pivot field's type (Row, Column, Data, Page) alongside its DisplayName in the output file.
// Common Searches: Aspose.Cells C# how to export pivot table field names to a text file | iterate over pivot table base fields and write display names with Aspose.Cells .NET | save pivot field metadata to a .txt file using Aspose.Cells library | log all pivot table fields from a workbook in C# Aspose.Cells example | write pivot table field display names to file programmatically in .NET
// Tags: Aspose.Cells write pivot field names to txt | C# enumerate pivot table base fields Aspose.Cells | log pivot field display names .NET | export pivot metadata Aspose.Cells example | iterate pivot fields and save to file C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, adds a pivot table, iterates through the pivot table's BaseFields collection, writes each field's DisplayName to PivotFieldsLog.txt, and saves the workbook as PivotFieldsDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["B1"].Value = "Amount";
        worksheet.Cells["A2"].Value = "Food";
        worksheet.Cells["B2"].Value = 100;
        worksheet.Cells["A3"].Value = "Drink";
        worksheet.Cells["B3"].Value = 150;
        worksheet.Cells["A4"].Value = "Food";
        worksheet.Cells["B4"].Value = 200;

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add fields to the pivot table (row and data fields)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Path for the log file
        string logFilePath = "PivotFieldsLog.txt";

        // Iterate through all base pivot fields and write their display names to the log file
        using (StreamWriter writer = new StreamWriter(logFilePath))
        {
            foreach (PivotField field in pivotTable.BaseFields)
            {
                writer.WriteLine(field.DisplayName);
            }
        }

        // Save the workbook with the pivot table
        workbook.Save("PivotFieldsDemo.xlsx");
    }
}
