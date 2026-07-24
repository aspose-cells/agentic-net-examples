// Title: Save Workbook as CSV with Native Data Types using Aspose.Cells for .NET
// Description: Demonstrates how to export an Aspose.Cells workbook to a CSV file while keeping each cell's original data type. The example creates a mixed‑type sheet, configures TxtSaveOptions (comma separator, minimal quoting, no quote prefix), and saves the result as output.csv.
// Keywords: Aspose.Cells CSV export | preserve cell data types | TxtSaveOptions C# | SaveFormat.Csv | Excel to CSV .NET | numeric values CSV Aspose | custom CSV delimiter | minimal quoting CSV
// Common Searches: Aspose.Cells export to CSV preserving data types | C# save Excel as CSV without converting numbers to text | TxtSaveOptions separator and quoting options | How to keep numeric cells numeric in CSV using Aspose.Cells | CSV export settings Aspose.Cells .NET
// Developer Intent: Export a workbook to CSV while retaining the original data types of all cells.
// Use Cases: Generate CSV reports for analytics pipelines that require numeric columns to remain numeric. | Create data exchange files for database imports where type fidelity is critical. | Produce lightweight CSV snapshots of mixed‑type worksheets with custom delimiters and only necessary quoting.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as TSV while preserving native cell types. | Show how to force all fields to be quoted in a CSV export using TxtSaveOptions. | Explain the role of the PreserveString property when exporting CSV files with Aspose.Cells.

using System;
using Aspose.Cells;
using System.Text;

// Demonstrates how to export an Aspose.Cells workbook to a CSV file while keeping each cell's original data type. The example creates a mixed‑type sheet, configures TxtSaveOptions (comma separator, minimal quoting, no quote prefix), and saves the result as output.csv.
class SaveWorkbookAsCsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with mixed data types
        worksheet.Cells["A1"].PutValue("Name");   // string header
        worksheet.Cells["B1"].PutValue("Age");    // string header
        worksheet.Cells["A2"].PutValue("John");   // string value
        worksheet.Cells["B2"].PutValue(30);       // numeric value
        worksheet.Cells["A3"].PutValue("Jane");   // string value
        worksheet.Cells["B3"].PutValue(25);       // numeric value

        // The Cells collection preserves native data types by default.
        // No need to change PreserveString (default is false).

        // Configure CSV save options
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            Separator = ',',          // Use comma as delimiter
            AlwaysQuoted = false,     // Quote only when necessary
            ExportQuotePrefix = false // Do not export leading quote characters
        };

        // Save the workbook as CSV while keeping original data types
        workbook.Save("output.csv", csvOptions);
    }
}
