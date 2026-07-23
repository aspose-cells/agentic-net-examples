// Title: Update built‑in and custom document properties, default style, and a cell value in an Excel workbook with Aspose.Cells for .NET and save as XLSX
// Description: C# example that loads an existing workbook, changes the Author built‑in property, adds a boolean custom property called Reviewed, sets the workbook's default font to Calibri 12, writes "Modified" to cell A1, and saves the file as a new XLSX preserving all changes using Aspose.Cells.
// Keywords: Aspose.Cells C# modify document properties | add custom property Aspose.Cells | change default font Aspose.Cells | update cell value Aspose.Cells | save workbook as XLSX Aspose.Cells
// Common Searches: Aspose.Cells change Author property .NET | how to add custom document property in Excel with Aspose.Cells | set default font for entire workbook using Aspose.Cells | write value to cell A1 with Aspose.Cells C# | save modified Excel file as XLSX using Aspose.Cells
// Developer Intent: Load an existing Excel file, modify its metadata, styling, and a cell, then save the workbook as an XLSX file with all changes applied.
// Use Cases: Standardize author metadata before publishing a report. | Flag a workbook as reviewed by adding a custom boolean property. | Apply a consistent default font across all worksheets after generation. | Programmatically update a specific cell while preserving document properties.
// AI Prompts: Generate C# code with Aspose.Cells to change the Author built‑in property and add a custom boolean property named Reviewed. | Show how to set Calibri 12 as the default font for an entire workbook and write "Modified" into cell A1 using Aspose.Cells for .NET. | Explain the steps to save a workbook after modifying properties, style, and cell content, ensuring all changes are retained in XLSX format with Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that loads an existing workbook, changes the Author built‑in property, adds a boolean custom property called Reviewed, sets the workbook's default font to Calibri 12, writes "Modified" to cell A1, and saves the file as a new XLSX preserving all changes using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load an existing workbook from disk
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Modify built‑in document property
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";

        // Add a custom document property
        workbook.CustomDocumentProperties.Add("Reviewed", true);

        // Change default style (font name and size)
        workbook.DefaultStyle.Font.Name = "Calibri";
        workbook.DefaultStyle.Font.Size = 12;

        // Example cell modification
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Modified");

        // Save the workbook preserving all changes in XLSX format
        string outputFile = "output.xlsx";
        workbook.Save(outputFile, SaveFormat.Xlsx);
    }
}
