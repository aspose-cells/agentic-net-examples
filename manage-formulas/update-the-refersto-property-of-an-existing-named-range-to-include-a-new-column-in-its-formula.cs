// Title: Expand an existing named range by one column using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, retrieves a named range, parses its RefersTo address, extends the range one column to the right, updates the RefersTo property, and saves the file. | Show how to programmatically modify the RefersTo string of a named range to include an additional column in a .NET application using Aspose.Cells.
// Common Searches: aspnet expand named range column Aspose.Cells C# example | update RefersTo property of a named range programmatically Aspose.Cells | add extra column to existing Excel named range using Aspose.Cells for .NET | C# code to change named range boundaries in an Excel workbook with Aspose.Cells | how to parse and modify RefersTo address of a named range in Aspose.Cells
// Tags: extend named range column Aspose.Cells | modify RefersTo address C# | named range boundary update .NET | Aspose.Cells range expansion example | Excel named range manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a workbook, locates a named range, parses its RefersTo address to determine current limits, calculates a new end column one column to the right, builds a new RefersTo string, assigns it back to the named range, and saves the updated workbook.
class UpdateNamedRange
{
    // Helper to convert zero‑based column index to Excel column letter (e.g., 0 -> "A")
    static string ColumnIndexToName(int index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string result = "";
        while (index >= 0)
        {
            result = letters[index % 26] + result;
            index = index / 26 - 1;
        }
        return result;
    }

    static void Main()
    {
        // Paths – adjust as needed
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Name of the existing named range to modify
        string namedRangeName = "MyRange";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range object
            Name namedRange = workbook.Worksheets.Names[namedRangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{namedRangeName}\" not found.");
                return;
            }

            // Parse the RefersTo string to obtain sheet name and cell addresses
            // Example RefersTo: Sheet1!$A$1:$C$5
            string refersTo = namedRange.RefersTo;
            int exclPos = refersTo.IndexOf('!');
            if (exclPos < 0)
            {
                Console.WriteLine("Invalid RefersTo format.");
                return;
            }

            string sheetName = refersTo.Substring(0, exclPos);
            string addressPart = refersTo.Substring(exclPos + 1); // "$A$1:$C$5"
            string[] cells = addressPart.Split(':');
            if (cells.Length != 2)
            {
                Console.WriteLine("Invalid address format in RefersTo.");
                return;
            }

            // Get the worksheet that contains the named range
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                return;
            }

            // Create a Range object from the start and end cells
            Aspose.Cells.Range range = sheet.Cells.CreateRange(cells[0], cells[1]);

            // Determine current boundaries
            int firstRow = range.FirstRow;
            int firstColumn = range.FirstColumn;
            int lastRow = range.RowCount > 0 ? firstRow + range.RowCount - 1 : firstRow;
            int lastColumn = range.ColumnCount > 0 ? firstColumn + range.ColumnCount - 1 : firstColumn;

            // Extend the range by one column to the right
            int newLastColumn = lastColumn + 1;

            // Build the new RefersTo address (e.g., Sheet1!$A$1:$D$10)
            string startCell = $"${ColumnIndexToName(firstColumn)}${firstRow + 1}";
            string endCell = $"${ColumnIndexToName(newLastColumn)}${lastRow + 1}";
            string newRefersTo = $"{sheet.Name}!{startCell}:{endCell}";

            // Update the RefersTo property
            namedRange.RefersTo = newRefersTo;

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Named range \"{namedRangeName}\" updated to \"{newRefersTo}\" and saved to {outputPath}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
