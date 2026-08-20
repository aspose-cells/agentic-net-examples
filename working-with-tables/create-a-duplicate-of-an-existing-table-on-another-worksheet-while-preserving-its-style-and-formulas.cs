// Title: Copy an Excel table to a new worksheet with styles and formulas – Aspose.Cells C# example
// Description: Loads a workbook, selects a table range, adds a destination sheet, and copies the range using PasteOptions.KeepOldTables so the table definition, formatting, and formulas are retained before saving.
// Keywords: Aspose.Cells copy table C# | duplicate Excel table Aspose | preserve table formatting Aspose.Cells | PasteOptions KeepOldTables | copy range with formulas .NET
// Common Searches: Aspose.Cells copy table to another sheet | C# duplicate Excel table preserving styles | PasteOptions KeepOldTables example | how to retain formulas when copying a table in Aspose.Cells | clone Excel table programmatically C#
// Developer Intent: Replicate an existing Excel table on a different worksheet while keeping its style, formulas, and table metadata intact.
// Use Cases: Create a styled snapshot of a data table on a summary sheet without breaking calculations. | Back up a worksheet‑level table on a separate tab for version control. | Generate a working copy of a table for further data manipulation while preserving original formulas.
// AI Prompts: Write C# code with Aspose.Cells that copies a table from Sheet1 to a new sheet, preserving formatting and formulas. | Explain the effect of PasteOptions.KeepOldTables when duplicating a table range in Aspose.Cells. | Suggest how to programmatically locate a table’s boundaries and clone it to another worksheet while retaining all metadata.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Loads a workbook, selects a table range, adds a destination sheet, and copies the range using PasteOptions.KeepOldTables so the table definition, formatting, and formulas are retained before saving.
class DuplicateTable
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the original table
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Get the worksheet where the original table resides
            Worksheet sourceSheet = sourceWorkbook.Worksheets["Sheet1"]; // adjust name as needed

            // Define the range that represents the table (including header, data and style)
            // Example range: A1:C5 – modify to match your actual table range
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

            // Add a new worksheet that will hold the duplicated table
            Worksheet destinationSheet = sourceWorkbook.Worksheets.Add("DuplicatedTable");

            // Define the destination range where the table will be copied.
            // Using the same size as the source ensures the table structure is preserved.
            AsposeRange destinationRange = destinationSheet.Cells.CreateRange("A1:C5");

            // Configure paste options:
            // KeepOldTables = true ensures that the copied range retains its table definition.
            PasteOptions pasteOptions = new PasteOptions
            {
                KeepOldTables = true
            };

            // Perform the copy – this duplicates the table with its styles and formulas.
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the workbook with the duplicated table.
            sourceWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
