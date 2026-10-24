// Title: Export a named range to a new workbook while preserving formatting with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export a specific named range to a separate Excel file preserving cell formatting | How to copy values and styles of a named range to a new workbook using Aspose.Cells for .NET | C# extract MyNamedRange from SourceWorkbook.xlsx and save as a new workbook with original formatting | Export named range to another workbook with Aspose.Cells while keeping cell styles | Copy a named range to a new worksheet preserving formatting in Aspose.Cells C#
// Tags: export named range Aspose.Cells C# | copy cell styles Aspose.Cells | named range to new workbook Aspose.Cells | preserve formatting when exporting range Aspose.Cells | CellArea.CreateCellArea usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads a source workbook, retrieves a named range, determines its cell area, creates a new workbook, copies each cell's value and style to the destination sheet, and saves the result, preserving the original formatting.
class ExportNamedRange
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string destPath = "ExportedRange.xlsx";
            const string rangeName = "MyNamedRange";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook containing the named range
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Retrieve the named range by its name
            Name namedRange = sourceWorkbook.Worksheets.Names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Parse the range reference (e.g., Sheet1!$A$1:$C$3) to a CellArea
            // Use the overload that accepts a name and a reference string
            CellArea rangeArea = CellArea.CreateCellArea(string.Empty, namedRange.RefersTo);
            int startRow = rangeArea.StartRow;
            int startColumn = rangeArea.StartColumn;
            int rowCount = rangeArea.EndRow - rangeArea.StartRow + 1;
            int columnCount = rangeArea.EndColumn - rangeArea.StartColumn + 1;

            // Get the worksheet that the named range refers to
            Worksheet sourceSheet = sourceWorkbook.Worksheets[namedRange.SheetIndex];

            // Create a new workbook to hold the exported range
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "ExportedRange";

            // Copy each cell's value and style to the destination worksheet
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell srcCell = sourceSheet.Cells[startRow + i, startColumn + j];
                    Cell destCell = destSheet.Cells[i, j];

                    // Copy value
                    destCell.PutValue(srcCell.Value);

                    // Copy style (formatting)
                    Style srcStyle = srcCell.GetStyle();
                    destCell.SetStyle(srcStyle);
                }
            }

            // Save the new workbook containing the exported range
            destWorkbook.Save(destPath);
            Console.WriteLine($"Exported range saved to \"{destPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
