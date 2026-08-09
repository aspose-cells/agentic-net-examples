// Title: Export All Worksheets to a Single CSV with Custom Delimiter and UTF‑8 Encoding using Aspose.Cells for .NET (C#)
// Description: Creates a workbook with two sheets, fills them with sample data, configures TxtSaveOptions for CSV (semicolon separator, UTF‑8 encoding, ExportAllSheets=true) and saves the entire workbook as one CSV file.
// Keywords: Aspose.Cells | C# | .NET | CSV export | custom delimiter | semicolon separator | UTF-8 encoding | ExportAllSheets | TxtSaveOptions | multiple worksheets to single CSV
// Common Searches: Aspose.Cells export all sheets to one CSV | C# save workbook as CSV with semicolon delimiter | set UTF-8 encoding for CSV export Aspose.Cells | how to use TxtSaveOptions for CSV in .NET | export multiple worksheets to single CSV file
// Developer Intent: Generate a single CSV file that contains data from every worksheet in a workbook, using a semicolon as the field separator and UTF‑8 character encoding.
// Use Cases: Consolidate related worksheet data into one CSV for bulk import into analytics platforms. | Produce CSV reports that comply with European locale standards (semicolon delimiter) while preserving Unicode characters. | Automate UTF‑8 encoded CSV generation for downstream web services or APIs.
// AI Prompts: Show how to export each worksheet to its own CSV file while keeping the semicolon delimiter and UTF‑8 encoding. | Provide an example that uses TxtSaveOptions to create a tab‑delimited CSV with ISO‑8859‑1 encoding. | Explain the impact of setting ExportAllSheets to false and how to select specific worksheets for CSV export.

using System;
using System.Text;
using Aspose.Cells;

// Creates a workbook with two sheets, fills them with sample data, configures TxtSaveOptions for CSV (semicolon separator, UTF‑8 encoding, ExportAllSheets=true) and saves the entire workbook as one CSV file.
class ExportWorksheetsToCsv
{
    static void Main()
    {
        // Create a new workbook and add a second worksheet
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");

        // Populate first worksheet with sample data
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Cells["A1"].PutValue("Name");
        sheet1.Cells["B1"].PutValue("Age");
        sheet1.Cells["A2"].PutValue("John");
        sheet1.Cells["B2"].PutValue(30);
        sheet1.Cells["A3"].PutValue("Jane");
        sheet1.Cells["B3"].PutValue(25);

        // Populate second worksheet with sample data
        Worksheet sheet2 = workbook.Worksheets[1];
        sheet2.Cells["A1"].PutValue("Product");
        sheet2.Cells["B1"].PutValue("Price");
        sheet2.Cells["A2"].PutValue("Apple");
        sheet2.Cells["B2"].PutValue(1.5);
        sheet2.Cells["A3"].PutValue("Banana");
        sheet2.Cells["B3"].PutValue(0.75);

        // Configure TxtSaveOptions for CSV export
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.Separator = ';' ;               // Custom delimiter
        saveOptions.Encoding = Encoding.UTF8;       // UTF‑8 encoding
        saveOptions.ExportAllSheets = true;         // Export every worksheet

        // Save all worksheets to a single CSV file
        workbook.Save("AllSheetsOutput.csv", saveOptions);
    }
}
