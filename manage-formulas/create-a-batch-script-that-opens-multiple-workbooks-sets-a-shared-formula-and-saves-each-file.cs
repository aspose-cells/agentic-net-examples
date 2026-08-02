// Title: C# batch processing of Excel workbooks – set shared formula with Aspose.Cells
// Description: A C# console app that loops through a list of Excel file paths, loads each workbook (or creates a new one), fills column A with sequential numbers, applies a shared formula (e.g., =A1*2) to a specified range starting at B1, forces formula calculation, and saves the file. Includes error handling for missing or inaccessible files.
// Keywords: Aspose.Cells | C# | shared formula | batch update | Excel workbook | SetSharedFormula | calculate formulas | bulk processing | .NET | Excel automation
// Common Searches: Aspose.Cells set shared formula in multiple workbooks | C# batch apply formula to Excel files | How to use SetSharedFormula with Aspose.Cells | Bulk update Excel workbooks C# Aspose | Apply same formula to many worksheets programmatically
// Developer Intent: Apply one shared formula to a defined range across several Excel files and persist the changes using Aspose.Cells for .NET.
// Use Cases: Process a predefined list of file paths, populate A1:A20 with incremental values, set B1:B20 to a shared formula, calculate results, and overwrite each workbook. | Read workbook locations from a configuration file, database, or JSON array and perform the same shared‑formula operation on every entry. | Extend the loop to iterate through all worksheets in each workbook, applying the shared formula to the same range on each sheet. | Integrate the routine into a CI/CD pipeline to ensure newly generated reports contain the required calculations automatically.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch apply a shared formula to a list of Excel files, creating a new workbook when a file does not exist. | Show how to load workbook paths from a JSON configuration and set a shared formula on each workbook with robust error handling. | Provide an example that iterates over every worksheet in each workbook and applies the same shared formula range using Aspose.Cells. | Write a PowerShell wrapper that calls the compiled C# program to process Excel files in a directory tree.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchSharedFormula
{
    // A C# console app that loops through a list of Excel file paths, loads each workbook (or creates a new one), fills column A with sequential numbers, applies a shared formula (e.g., =A1*2) to a specified range starting at B1, forces formula calculation, and saves the file. Includes error handling for missing or inaccessible files.
    class Program
    {
        static void Main(string[] args)
        {
            // List of workbook file paths to process
            List<string> workbookPaths = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
                // Add more paths as needed
            };

            // Define the shared formula and the range size
            string sharedFormula = "=A1*2"; // Example: double the value in column A
            int rowsToPopulate = 20;        // Number of rows the shared formula will cover
            int columnsToPopulate = 1;      // Number of columns (only column B in this case)

            foreach (string path in workbookPaths)
            {
                try
                {
                    Workbook workbook;

                    // Load existing workbook if it exists; otherwise create a new one
                    if (File.Exists(path))
                    {
                        // Load without password (assumes file is not password‑protected)
                        workbook = new Workbook(path);
                    }
                    else
                    {
                        workbook = new Workbook(); // creates a default workbook with one worksheet
                    }

                    // Access the first worksheet (you can modify to target specific sheets)
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cells cells = worksheet.Cells;

                    // Ensure there are enough values in column A for the formula to work
                    for (int i = 0; i < rowsToPopulate; i++)
                    {
                        cells[i, 0].PutValue(i + 1); // Populate A1:A20 with 1,2,3,...
                    }

                    // Set the shared formula starting at B1
                    // Using the overload: SetSharedFormula(string, int, int)
                    cells["B1"].SetSharedFormula(sharedFormula, rowsToPopulate, columnsToPopulate);

                    // Calculate formulas so that results are stored
                    workbook.CalculateFormula();

                    // Save the modified workbook (overwrites the original file or creates a new one)
                    workbook.Save(path);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{path}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
