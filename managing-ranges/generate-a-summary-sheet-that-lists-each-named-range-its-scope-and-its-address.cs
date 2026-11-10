// Title: Create a summary worksheet that lists every named range, its scope, and address in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to export all defined names to a new worksheet | list named ranges with their scope and address using Aspose.Cells .NET | generate an Excel summary sheet of named ranges programmatically | retrieve the RefersTo address of a named range with Aspose.Cells in C# | save a workbook that includes a named‑range report using Aspose.Cells
// Tags: list defined names Aspose.Cells C# | write named range metadata worksheet | extract named range address Aspose.Cells | generate named range summary sheet .NET | auto-fit columns after inserting data Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one), adds a worksheet named "Summary", writes headers for Named Range, Scope, and Address, then iterates through workbook.Worksheets.Names. For each defined name it records the name text, assumes a "Workbook" scope, extracts the RefersTo address (removing a leading '='), and writes these values to the summary sheet. After auto‑fitting the columns and ensuring the output directory exists, the workbook is saved as "output.xlsx" with exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the input file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a new workbook with a default sheet
            }

            // Add a new worksheet for the summary and name it "Summary".
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write header titles.
            Cells cells = summarySheet.Cells;
            cells["A1"].PutValue("Named Range");
            cells["B1"].PutValue("Scope");
            cells["C1"].PutValue("Address");

            int currentRow = 1; // Start writing data from the second row.

            // Iterate through all defined names in the workbook.
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // Aspose.Cells older versions may not expose scope properties.
                // Default to "Workbook" scope.
                string scope = "Workbook";

                // Get the address the name refers to (remove leading '=' if present).
                string address = definedName.RefersTo;
                if (!string.IsNullOrEmpty(address) && address.StartsWith("="))
                {
                    address = address.Substring(1);
                }

                // Populate the summary sheet.
                cells[currentRow, 0].PutValue(definedName.Text); // Named range.
                cells[currentRow, 1].PutValue(scope);           // Scope.
                cells[currentRow, 2].PutValue(address);         // Address.

                currentRow++;
            }

            // Adjust column widths to fit content.
            summarySheet.AutoFitColumns();

            // Ensure the directory for the output file exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
