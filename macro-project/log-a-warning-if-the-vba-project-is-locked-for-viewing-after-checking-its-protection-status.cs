// Title: C# code to detect and log a warning when an Excel workbook’s VBA project is locked with Aspose.Cells
// Common Searches: how to check if an Excel VBA project is locked using Aspose.Cells in C# | log a warning in .NET when workbook VBA project protection is enabled | detect protected VBA macro in an .xlsx file with Aspose.Cells C# example | Aspose.Cells C# verify VBA project presence and protection status | C# handle missing VBA project in Excel workbook using Aspose.Cells
// Tags: Aspose.Cells VBA project protection detection | C# check Excel VBA project locked status | log warning for protected VBA macro | manage absent VBA project in workbook | IsProtected property usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, verifies that a VBA project exists, checks the IsProtected flag, and writes a warning to the console when the VBA project is locked, while also handling cases where no VBA project is present or errors occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook contains a VBA project
            if (workbook.VbaProject != null)
            {
                // Check if the VBA project is protected (locked)
                if (workbook.VbaProject.IsProtected)
                {
                    Console.WriteLine("Warning: The VBA project is locked.");
                }
                else
                {
                    Console.WriteLine("The VBA project is not locked.");
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }

            // Optionally save the workbook if modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
