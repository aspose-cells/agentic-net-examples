// Title: Save a PivotTable Workbook to XLSX with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a PivotTable from range A1:B4, configures row and data fields, and saves it as PivotTableDemo.xlsx using Aspose.Cells default save options.
// Keywords: Aspose.Cells C# pivot table | save workbook as xlsx | default save options Aspose.Cells | create pivot table .NET | export pivot table to xlsx | Aspose.Cells PivotTable example | C# Excel export Aspose
// Common Searches: Aspose.Cells add PivotTable C# | How to save workbook with PivotTable using Aspose.Cells | C# export PivotTable to XLSX | Aspose.Cells default save options example | Create and save PivotTable in .NET
// Developer Intent: Generate an Excel file that contains a configured PivotTable and write it to disk using the library’s default save behavior.
// Use Cases: Generate sales summary reports by building a PivotTable from raw data and exporting the workbook as XLSX for distribution. | Automate monthly inventory analysis by embedding a PivotTable in a generated workbook and saving the result for downstream processing. | Add an Excel export feature with a PivotTable to a web or desktop application that delivers ready‑to‑use XLSX files. | Create template‑driven financial dashboards that include PivotTables and are saved as XLSX workbooks.
// AI Prompts: Write C# code using Aspose.Cells to build a PivotTable from a data range and save the workbook as XLSX with default options. | Show how to add a column field to the existing PivotTable before saving the file. | Demonstrate exporting multiple worksheets, each containing a PivotTable, to separate XLSX files using Aspose.Cells. | Explain how to apply custom save settings such as password protection while preserving default PivotTable behavior.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds a PivotTable from range A1:B4, configures row and data fields, and saves it as PivotTableDemo.xlsx using Aspose.Cells default save options.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill sample data for the pivot table
        Cells cells = sheet.Cells;
        cells["A1"].Value = "Fruit";
        cells["B1"].Value = "Quantity";
        cells["A2"].Value = "Apple";
        cells["B2"].Value = 10;
        cells["A3"].Value = "Orange";
        cells["B3"].Value = 15;
        cells["A4"].Value = "Banana";
        cells["B4"].Value = 20;

        // Add a pivot table that uses the data range A1:B4 and places the table at D1
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: Fruit as row field, Quantity as data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Save the workbook to XLSX format using default save options
        workbook.Save("PivotTableDemo.xlsx");
    }
}
