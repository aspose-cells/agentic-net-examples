// Title: AutoFitRows after InsertRows in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, insert new rows, populate them, and then call sheet.AutoFitRows() (optionally sheet.AutoFitColumns()) to automatically adjust row heights so the added content aligns with existing rows before saving the file.
// Keywords: Aspose.Cells AutoFitRows C# | InsertRows Aspose.Cells .NET | adjust row height programmatically | auto‑fit rows after inserting rows | uniform row height Aspose.Cells
// Common Searches: Aspose.Cells AutoFitRows after InsertRows | C# auto fit rows after adding rows | how to resize rows to content in Aspose.Cells | auto‑fit worksheet rows .NET
// Developer Intent: Resize all worksheet rows to fit their content after new rows have been inserted.
// Use Cases: Generate a dynamic report, insert placeholder rows, fill them, and auto‑fit rows to keep the layout tidy. | Add monthly summary rows to an existing template and ensure proper row height before exporting to XLSX. | Programmatically build a spreadsheet, insert data rows on the fly, and apply AutoFitRows/AutoFitColumns for optimal display.
// AI Prompts: Provide C# code that inserts rows with Aspose.Cells, fills the cells, and then calls AutoFitRows to adjust heights. | Show an example of using sheet.AutoFitRows() after cells.InsertRows() and saving the workbook as XLSX. | Explain when and why to invoke AutoFitRows in Aspose.Cells after modifying worksheet data.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, insert new rows, populate them, and then call sheet.AutoFitRows() (optionally sheet.AutoFitColumns()) to automatically adjust row heights so the added content aligns with existing rows before saving the file.
class AutoFitRowsAfterInsertDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add some initial data
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Original Row 1");
        cells["A3"].PutValue("Original Row 2");

        // Insert two new rows at index 2 (zero‑based). Existing rows shift down.
        cells.InsertRows(2, 2);

        // Populate the newly inserted rows
        cells["A3"].PutValue("Inserted Row 1");
        cells["A4"].PutValue("Inserted Row 2");

        // Auto‑fit all rows to adjust their heights based on the new content
        sheet.AutoFitRows();

        // (Optional) Auto‑fit columns for better visibility
        sheet.AutoFitColumns();

        // Save the workbook
        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "AutoFitRowsAfterInsert.xlsx");
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine("Workbook saved to: " + outputPath);
    }
}
