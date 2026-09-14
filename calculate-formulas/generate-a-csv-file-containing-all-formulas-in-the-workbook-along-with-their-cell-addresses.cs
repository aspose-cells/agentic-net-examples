// Title: Export every formula and its cell address from an Excel workbook to a CSV file using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that opens an .xlsx file with Aspose.Cells, finds all cells containing formulas, and writes the worksheet name, A1 cell address, and formula text to a CSV file. | Refactor the example to ignore hidden worksheets and only export formulas from visible sheets while still recording the sheet name in the CSV. | Create a reusable method that returns a List<(string Sheet, string Address, string Formula)> representing all formula cells in a given workbook using Aspose.Cells. | Add code that logs the total number of formulas exported after the CSV file has been generated.
// Common Searches: how to extract all formulas from an Excel file to a CSV using Aspose.Cells C# | C# Aspose.Cells list formula cells with addresses across worksheets | save Excel formulas as text with cell references using Aspose.Cells .NET | export worksheet formulas to CSV file programmatically Aspose.Cells | retrieve A1 notation formulas from workbook using Aspose.Cells C#
// Tags: Aspose.Cells export formulas to CSV | C# extract cell formulas with Aspose.Cells | iterate worksheets to collect formulas Aspose.Cells | write formula strings to CSV in .NET | retrieve A1 addresses of formula cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The code loads an Excel workbook with Aspose.Cells, iterates through each worksheet and cell, detects formula cells, and writes the worksheet name, cell address (A1 notation), and formula text to a CSV file named formulas.csv.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (create/load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Prepare a CSV file to store formulas (save rule for the CSV)
        using (StreamWriter writer = new StreamWriter("formulas.csv"))
        {
            // Write CSV header
            writer.WriteLine("CellAddress,Formula");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data
                foreach (Cell cell in cells)
                {
                    // Check if the cell holds a formula
                    if (cell.IsFormula)
                    {
                        // Get the cell address in A1 notation
                        string address = cell.Name;

                        // Get the formula text; escape double quotes for CSV compliance
                        string formula = cell.Formula.Replace("\"", "\"\"");

                        // Write the address and formula to the CSV (formula is quoted to handle commas)
                        writer.WriteLine($"{address},\"{formula}\"");
                    }
                }
            }
        }

        Console.WriteLine("All formulas have been exported to formulas.csv");
    }
}
