// Title: Copy a Worksheet and Freeze Header Row with Aspose.Cells for .NET (C#)
// Description: Loads or creates a source workbook, creates an empty destination workbook, copies the first worksheet using Worksheets.AddCopy, applies FreezePanes to lock the top row, and saves the result. Demonstrates worksheet duplication and header freezing in Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | copy worksheet | AddCopy | FreezePanes | freeze top row | duplicate sheet | Excel automation | workbook merge | header freeze
// Common Searches: Aspose.Cells copy worksheet between workbooks C# | How to freeze the first row after copying a sheet with Aspose.Cells | AddCopy method example Aspose.Cells .NET | FreezePanes on copied worksheet Aspose.Cells | C# copy Excel sheet and lock header row
// Developer Intent: Duplicate a worksheet from one workbook to another and keep the header row visible by applying FreezePanes.
// Use Cases: Create a reporting file by copying a template sheet and freezing its header for scrolling users. | Consolidate multiple source workbooks into a single file, copying each sheet and applying FreezePanes for easy navigation. | Generate a summary workbook that reuses a master data sheet while keeping column titles fixed.
// AI Prompts: Generate C# code using Aspose.Cells to copy the first worksheet from source.xlsx to a new workbook and freeze its top row. | Explain the role of Worksheets.AddCopy and FreezePanes in Aspose.Cells and suggest alternative approaches for the same outcome. | Provide a step‑by‑step guide to copy several worksheets from different workbooks into one workbook, applying FreezePanes to each copied sheet.

using System;
using System.IO;
using Aspose.Cells;

// Loads or creates a source workbook, creates an empty destination workbook, copies the first worksheet using Worksheets.AddCopy, applies FreezePanes to lock the top row, and saves the result. Demonstrates worksheet duplication and header freezing in Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the source workbook exists; create a simple one if missing
            Workbook sourceWorkbook;
            if (File.Exists(sourcePath))
            {
                sourceWorkbook = new Workbook(sourcePath);
            }
            else
            {
                sourceWorkbook = new Workbook();
                Worksheet ws = sourceWorkbook.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["A1"].PutValue("Sample Data");
                sourceWorkbook.Save(sourcePath);
            }

            // Create a new (empty) destination workbook and clear the default sheet
            Workbook destinationWorkbook = new Workbook();
            destinationWorkbook.Worksheets.Clear();

            // Copy the first worksheet from the source workbook into the destination workbook
            // AddCopy(string) copies the worksheet by its name and returns the new index
            int copiedIndex = destinationWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[0].Name);
            Worksheet copiedSheet = destinationWorkbook.Worksheets[copiedIndex];

            // Freeze the header row (first row) in the copied worksheet
            copiedSheet.FreezePanes(1, 0, 1, 0);

            // Save the resulting workbook
            destinationWorkbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
