// Title: Import a cell range from one Excel workbook into a named range of another workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load a source .xlsx file, extract a specific range, erase the target named range, and paste the values into it with Aspose.Cells in C#. | Programmatically replace the contents of a defined name in an existing workbook with data from another workbook using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# copy range from one workbook to a named range in another workbook | Overwrite data in an Excel defined name using Aspose.Cells .NET | Clear existing cells in a named range before importing values with Aspose.Cells | How to transfer A1:C10 from source workbook to MyNamedRange in target workbook using C# | Aspose.Cells example for updating a named range with external workbook data
// Tags: copy source range into target named range Aspose.Cells | erase existing cells in a defined name Aspose.Cells | load multiple workbooks and move data C# | populate named range from another workbook using Aspose.Cells | overwrite named range values using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source and a target workbook, extracts a defined cell block from the source, locates a named range in the target, clears any existing data in that range, copies the source values into it, and saves the updated target workbook.
class ImportToNamedRange
{
    static void Main()
    {
        try
        {
            // Define file paths
            string sourcePath = "SourceWorkbook.xlsx";
            string targetPath = "TargetWorkbook.xlsx";

            // Verify that the required files exist
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");
            if (!File.Exists(targetPath))
                throw new FileNotFoundException($"Target file not found: {targetPath}");

            // Load the source workbook containing the data to import
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Load the target workbook that contains the named range
            Workbook targetWorkbook = new Workbook(targetPath);

            // ------------------------------------------------------------
            // 1. Retrieve the source range (adjust the sheet name and range as needed)
            // ------------------------------------------------------------
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0]; // first sheet
            // Example source range A1:C10 – change to your actual range
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:C10");

            // ------------------------------------------------------------
            // 2. Locate the named range in the target workbook
            // ------------------------------------------------------------
            // Get the range object that corresponds to the defined name
            Aspose.Cells.Range namedRange = targetWorkbook.Worksheets.GetRangeByName("MyNamedRange");
            if (namedRange == null)
                throw new InvalidOperationException("Named range 'MyNamedRange' does not exist in the target workbook.");

            Worksheet targetSheet = namedRange.Worksheet;

            // Determine the top‑left cell of the named range
            int targetStartRow = namedRange.FirstRow;
            int targetStartColumn = namedRange.FirstColumn;

            // ------------------------------------------------------------
            // 3. Clear existing data inside the named range (optional but safe)
            // ------------------------------------------------------------
            targetSheet.Cells.ClearRange(targetStartRow, targetStartColumn,
                                         namedRange.RowCount, namedRange.ColumnCount);

            // ------------------------------------------------------------
            // 4. Copy the source values into the named range
            // ------------------------------------------------------------
            // Create a destination range that matches the size of the source range
            Aspose.Cells.Range destinationRange = targetSheet.Cells.CreateRange(
                                                    targetStartRow,
                                                    targetStartColumn,
                                                    sourceRange.RowCount,
                                                    sourceRange.ColumnCount);

            // Perform the copy – this overwrites any existing data within the destination range
            destinationRange.Copy(sourceRange);

            // ------------------------------------------------------------
            // 5. Save the modified target workbook
            // ------------------------------------------------------------
            string outputPath = "TargetWorkbook_Updated.xlsx";
            targetWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
