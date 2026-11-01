// Title: Enable the 1904 date system for legacy formulas in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# set workbook to use 1904 date system for legacy formulas | How to configure Excel 1904 date system with Aspose.Cells API in C# | C# example for enabling 1904 date system in an Aspose.Cells workbook | Configure Aspose.Cells calculation engine to use 1904 date system for older Excel files | Legacy Excel date handling with Aspose.Cells 1904 date system in C#
// Tags: Aspose.Cells workbook Settings.Date1904 | 1904 date system configuration Aspose.Cells | legacy Excel date formulas Aspose.Cells | C# enable Excel 1904 date system | Aspose.Cells calculation engine date system

using System;
using Aspose.Cells;
using System.IO;

namespace AsposeCellsExample
{
    // The example creates a new Aspose.Cells Workbook, activates the 1904 date system by setting workbook.Settings.Date1904 to true, saves the workbook as "output.xlsx", and includes basic exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                // If loading: check file existence first
                // string inputPath = "input.xlsx";
                // if (File.Exists(inputPath))
                //     workbook = new Workbook(inputPath);
                // else
                //     throw new FileNotFoundException($"Input file not found: {inputPath}");

                Workbook workbook = new Workbook();

                // Enable the 1904 date system for legacy date‑dependent formulas
                workbook.Settings.Date1904 = true;

                // Save the workbook to a file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
