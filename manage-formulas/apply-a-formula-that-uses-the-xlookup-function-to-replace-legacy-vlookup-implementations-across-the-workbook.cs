// Title: Programmatically replace VLOOKUP with XLOOKUP in every worksheet of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to convert all VLOOKUP formulas to XLOOKUP in an Excel file with Aspose.Cells | How to update legacy VLOOKUP formulas to XLOOKUP across multiple sheets using Aspose.Cells .NET | Automate Excel formula migration from VLOOKUP to XLOOKUP in .NET applications | Replace VLOOKUP with XLOOKUP in a workbook programmatically using Aspose.Cells API | Iterate through worksheets and change VLOOKUP to XLOOKUP in C# with Aspose.Cells
// Tags: Aspose.Cells VLOOKUP to XLOOKUP conversion | C# iterate workbook cells modify formulas | generate XLOOKUP formula from VLOOKUP arguments | Excel formula migration using Aspose.Cells | replace legacy lookup formulas .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, walks through each worksheet and every used cell, detects VLOOKUP formulas, parses their arguments, builds an equivalent XLOOKUP expression with correct lookup and return ranges and match mode, substitutes the original formula, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the sheet; skip if the sheet is empty
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange == null) continue;

                // Loop through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Identify cells that contain a VLOOKUP formula
                    if (cell.IsFormula && cell.Formula.Contains("VLOOKUP"))
                    {
                        string oldFormula = cell.Formula; // e.g., =VLOOKUP(A2,Sheet2!$A$1:$B$100,2,FALSE)

                        // Extract the arguments inside VLOOKUP(...)
                        int startIdx = oldFormula.IndexOf('(');
                        int endIdx = oldFormula.LastIndexOf(')');
                        if (startIdx < 0 || endIdx < 0 || endIdx <= startIdx + 1) continue;

                        string inner = oldFormula.Substring(startIdx + 1, endIdx - startIdx - 1);
                        string[] parts = inner.Split(',');

                        // Ensure we have at least the four required arguments
                        if (parts.Length >= 4)
                        {
                            string lookupValue = parts[0].Trim();          // A2
                            string tableArray = parts[1].Trim();           // Sheet2!$A$1:$B$100
                            string colIndexStr = parts[2].Trim();          // 2
                            string rangeLookup = parts[3].Trim();          // FALSE

                            // Split the table array into start and end addresses
                            string[] rangeParts = tableArray.Split(':');
                            if (rangeParts.Length == 2)
                            {
                                string startAddr = rangeParts[0];
                                string endAddr = rangeParts[1];

                                // Determine the first column (lookup column) of the table array
                                string startCol = GetColumnLetter(startAddr);
                                string endCol = GetColumnLetter(endAddr);

                                // Build the lookup column range (first column of the table)
                                string lookupColumnRange = $"{ReplaceColumn(startAddr, startCol)}:{ReplaceColumn(endAddr, startCol)}";

                                // Parse the column index to know which column to return
                                if (int.TryParse(colIndexStr, out int colIndex) && colIndex > 0)
                                {
                                    // Calculate the return column letter based on the index
                                    string returnCol = ShiftColumn(startCol, colIndex - 1);
                                    string returnColumnRange = $"{ReplaceColumn(startAddr, returnCol)}:{ReplaceColumn(endAddr, returnCol)}";

                                    // Determine match mode: 0 = exact match (FALSE), -1 = approximate (TRUE)
                                    string matchMode = (rangeLookup.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
                                                        rangeLookup.Equals("0", StringComparison.OrdinalIgnoreCase))
                                                        ? "0" : "-1";

                                    // Construct the XLOOKUP formula
                                    string newFormula = $"=XLOOKUP({lookupValue},{lookupColumnRange},{returnColumnRange},, {matchMode})";

                                    // Apply the new XLOOKUP formula to the cell
                                    cell.Formula = newFormula;
                                }
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper: Extract column letters from an address (e.g., Sheet2!$A$1 -> A)
    static string GetColumnLetter(string address)
    {
        // Remove sheet name if present
        int excl = address.IndexOf('!');
        if (excl >= 0) address = address.Substring(excl + 1);

        // Find the first '$' and take characters after it until a digit appears
        int dollar = address.IndexOf('$');
        if (dollar >= 0) address = address.Substring(dollar + 1);
        int i = 0;
        while (i < address.Length && Char.IsLetter(address[i]))
            i++;
        return address.Substring(0, i);
    }

    // Helper: Replace the column part of an address with a new column letter
    static string ReplaceColumn(string address, string newCol)
    {
        // Preserve sheet name and row number
        int excl = address.IndexOf('!');
        string sheetPart = excl >= 0 ? address.Substring(0, excl + 1) : "";
        int rowStart = address.IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
        string rowPart = rowStart >= 0 ? address.Substring(rowStart) : "";
        return $"{sheetPart}${newCol}${rowPart}";
    }

    // Helper: Shift a column letter by an offset (e.g., A + 1 => B)
    static string ShiftColumn(string col, int offset)
    {
        int number = 0;
        foreach (char c in col.ToUpperInvariant())
        {
            number = number * 26 + (c - 'A' + 1);
        }
        number += offset;
        string result = "";
        while (number > 0)
        {
            int rem = (number - 1) % 26;
            result = (char)('A' + rem) + result;
            number = (number - 1) / 26;
        }
        return result;
    }
}
