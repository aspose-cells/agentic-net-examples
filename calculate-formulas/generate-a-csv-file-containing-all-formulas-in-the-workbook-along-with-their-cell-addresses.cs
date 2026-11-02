// Title: Export all formulas with their cell addresses from an Excel workbook to a CSV file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to list formulas with cell references | generate CSV containing Excel formulas using Aspose.Cells .NET | save workbook formulas to a CSV file in C# | extract all formula strings from an .xlsx file with Aspose.Cells | create a CSV of cell addresses and formulas from a workbook in .NET
// Tags: Aspose.Cells export formulas to CSV | retrieve cell formulas C# Aspose.Cells | write formula list CSV .NET | iterate worksheet cells Aspose.Cells | extract Excel formulas programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and cell, detects cells that contain formulas, escapes double quotes for CSV compliance, and writes the cell address together with its formula to a UTF‑8 CSV file with a header. It also includes error handling for missing files and password‑protected workbooks.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Path to the output CSV file
        string csvPath = "formulas.csv";

        // Verify that the input workbook exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The workbook file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (no password)
            Workbook workbook = new Workbook(workbookPath);

            // Create a StreamWriter for the CSV file
            using (StreamWriter writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8))
            {
                // Write CSV header
                writer.WriteLine("CellAddress,Formula");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all cells in the worksheet
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Check if the cell contains a formula
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Escape double quotes in the formula for CSV compliance
                            string escapedFormula = cell.Formula.Replace("\"", "\"\"");

                            // Write the cell address and its formula to the CSV
                            writer.WriteLine($"{cell.Name},\"{escapedFormula}\"");
                        }
                    }
                }
            }

            Console.WriteLine("CSV file with formulas has been created at: " + Path.GetFullPath(csvPath));
        }
        catch (CellsException ex) when (ex.Message.Contains("password"))
        {
            Console.WriteLine("Error: The workbook is password protected. Please provide a valid password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
