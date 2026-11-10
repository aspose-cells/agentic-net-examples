// Title: How to programmatically extend an Excel named range by one column using Aspose.Cells for .NET
// Common Searches: C# code to add an extra column to an existing named range in an Excel workbook with Aspose.Cells | Aspose.Cells .NET update RefersTo property to include a new column | extend named range range address dynamically in C# using Aspose.Cells | parse and modify a named range RefersTo string in Aspose.Cells for .NET | calculate next column index and adjust a named range in Excel via Aspose.Cells
// Tags: update named range RefersTo Aspose.Cells | extend named range column C# | calculate next column index Aspose.Cells | modify Excel named range formula .NET | dynamic range address adjustment Aspose.Cells | parse RefersTo address C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, retrieves a named range, parses its RefersTo address to find the current end column, computes the next column index, builds a new RefersTo string that adds one column, assigns the updated RefersTo back to the named range, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range (replace "MyRange" with your actual name)
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Current RefersTo string, e.g. Sheet1!$A$1:$C$10
            string currentRefersTo = namedRange.RefersTo;

            // Split into sheet name and range part
            int exclPos = currentRefersTo.IndexOf('!');
            if (exclPos < 0)
            {
                Console.WriteLine("Invalid RefersTo format.");
                return;
            }

            string sheetName = currentRefersTo.Substring(0, exclPos);
            string rangePart = currentRefersTo.Substring(exclPos + 1); // "$A$1:$C$10"

            // Separate start and end cells
            string[] cells = rangePart.Split(':');
            if (cells.Length != 2)
            {
                Console.WriteLine("RefersTo does not contain a valid range.");
                return;
            }

            string startCell = cells[0]; // "$A$1"
            string endCell = cells[1];   // "$C$10"

            // Determine the column index of the current end cell
            string endColName = GetColumnName(endCell);               // "C"
            int endColIndex = ColumnNameToIndex(endColName);          // 2 (0‑based)

            // Extend the range by one column (e.g., to column D)
            int newEndColIndex = endColIndex + 1;
            string newEndColName = IndexToColumnName(newEndColIndex); // "D"

            // Preserve the row part of the end cell
            string endRowName = GetRowName(endCell); // "10"

            // Build the new end cell string (keep $ signs)
            string newEndCell = $"${newEndColName}${endRowName}";

            // Construct the updated RefersTo formula
            string newRefersTo = $"{sheetName}!{startCell}:{newEndCell}";

            // Apply the updated RefersTo to the named range
            namedRange.RefersTo = newRefersTo;

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Extract column letters from a cell reference like "$A$1"
    private static string GetColumnName(string cellReference)
    {
        int start = cellReference.IndexOf('$') + 1;
        int end = cellReference.IndexOf('$', start);
        if (end < 0) end = cellReference.Length;
        return cellReference.Substring(start, end - start);
    }

    // Extract row number from a cell reference like "$A$1"
    private static string GetRowName(string cellReference)
    {
        int lastDollar = cellReference.LastIndexOf('$');
        return cellReference.Substring(lastDollar + 1);
    }

    // Convert column name (e.g., "C") to zero‑based index
    private static int ColumnNameToIndex(string columnName)
    {
        int index = 0;
        foreach (char c in columnName.ToUpper())
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index - 1;
    }

    // Convert zero‑based column index to column name (e.g., 3 -> "D")
    private static string IndexToColumnName(int index)
    {
        index++; // make it 1‑based
        string name = string.Empty;
        while (index > 0)
        {
            int remainder = (index - 1) % 26;
            name = (char)('A' + remainder) + name;
            index = (index - remainder) / 26;
        }
        return name;
    }
}
