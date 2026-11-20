// Title: Load a password‑protected Excel workbook, select a worksheet by name, and write a value to cell B2 using Aspose.Cells for .NET
// Common Searches: C# open encrypted Excel file with Aspose.Cells and update a cell | Aspose.Cells .NET retrieve worksheet "Sheet1" and set cell B2 value | how to prompt for password when loading an Excel workbook with Aspose.Cells | write data to a specific cell and save the workbook using Aspose.Cells C# example | handle load exceptions while modifying an Excel file with Aspose.Cells
// Tags: load password protected Excel Aspose.Cells | select worksheet by name Aspose.Cells | write value to cell B2 Aspose.Cells | save modified workbook Aspose.Cells | handle workbook load exceptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program verifies the input file, loads an Excel workbook with Aspose.Cells (prompting for a password if encrypted), selects the worksheet named "Sheet1" (or the first sheet as fallback), obtains cell B2, writes "Hello Aspose!" to it, and saves the updated workbook to a new file while handling possible I/O and loading errors.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;

            // Try to load the workbook; if it is password‑protected, ask for the password
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (CellsException)
            {
                // Workbook is encrypted – request password from the user
                Console.Write("Workbook is password protected. Enter password: ");
                string password = Console.ReadLine();

                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };

                try
                {
                    workbook = new Workbook(inputPath, loadOptions);
                }
                catch (CellsException ce)
                {
                    Console.WriteLine($"Failed to open workbook with provided password: {ce.Message}");
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading workbook with password: {ex.Message}");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Access the desired worksheet (by name, fallback to first sheet)
            Worksheet worksheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

            // Obtain reference to cell B2 and write a value
            Cell targetCell = worksheet.Cells["B2"];
            targetCell.PutValue("Hello Aspose!");

            // Save the modified workbook
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
