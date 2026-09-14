// Title: Find cells whose formulas reference a specific named range using Aspose.Cells FindOptions in C#
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, configures FindOptions to search within formulas, and prints the addresses of all cells that contain a given named range. | Demonstrate how to loop through every formula cell that references a particular named range by using FindOptions with LookInType.Formulas and LookAtType.Contains. | Create a reusable method that accepts a workbook path and a named range, and returns a list of cell addresses where the range is referenced in formulas.
// Common Searches: asp.net find cells that reference a named range using aspose.cells FindOptions | c# search Excel formulas for a specific named range with Aspose.Cells | how to locate all formula cells containing a particular named range in a workbook using Aspose | using FindOptions to detect named range references in Excel formulas C#
// Tags: FindOptions search formulas Aspose.Cells | locate formula cells referencing named range C# | search Excel formulas for specific text Aspose | iterate cells with named range reference Aspose.Cells | lookup named range in formula cells using FindOptions

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file, sets FindOptions to look only in formulas with a Contains match, and iterates through every cell whose formula references the specified named range, outputting each cell's address. Optional saving of the workbook is shown.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string namedRange = "MyNamedRange";

            try
            {
                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: File '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure FindOptions to search within formulas and look for the named range text
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Formulas,      // Search only in formulas
                    LookAtType = LookAtType.Contains       // Match if the formula contains the text
                };

                // Find the first cell that contains the named range in its formula
                Cell cell = worksheet.Cells.Find(namedRange, null, findOptions);

                // Iterate through all matching cells
                while (cell != null)
                {
                    Console.WriteLine($"Cell {cell.Name} contains a formula referencing '{namedRange}'.");
                    // Find the next occurrence after the current cell
                    cell = worksheet.Cells.Find(namedRange, cell, findOptions);
                }

                // (Optional) Save the workbook if any modifications were made
                // workbook.Save("Output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
