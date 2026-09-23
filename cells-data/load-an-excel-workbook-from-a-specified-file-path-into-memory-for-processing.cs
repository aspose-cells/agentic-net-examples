// Title: Load an Excel workbook from a file path with Aspose.Cells in C# and create a new workbook if the file is missing
// AI Prompts: Write C# code that uses Aspose.Cells to open an Excel workbook from a specified path, and if the file does not exist, instantiate a new workbook and save it to the same location. | Show how to check for the existence of an Excel file before loading it with Aspose.Cells, then access the first worksheet after the workbook is loaded or created.
// Common Searches: aspocells open workbook from specific path c# | c# create new Excel workbook when file not found using Aspose.Cells | check file existence before loading Excel with Aspose.Cells | save newly created workbook to original path aspocells | access first worksheet after loading workbook aspocells c#
// Tags: load workbook from file Aspose.Cells C# | create workbook on missing file Aspose.Cells | save workbook to same location Aspose.Cells | verify Excel file existence before opening Aspose.Cells | access first worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // C# example that loads an Excel workbook from a given file path using Aspose.Cells; if the file is absent, it creates a new workbook, saves it to the same path, and then accesses the first worksheet.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = @"C:\Path\To\Your\Workbook.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(filePath))
                {
                    workbook = new Workbook(filePath);
                    Console.WriteLine("Workbook loaded from file.");
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                    Console.WriteLine("File not found. A new workbook has been created.");

                    // Optionally save the new workbook for future runs
                    try
                    {
                        workbook.Save(filePath);
                        Console.WriteLine($"New workbook saved to '{filePath}'.");
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine($"Failed to save new workbook: {saveEx.Message}");
                    }
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("Loaded worksheet name: " + sheet.Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook: " + ex.Message);
            }

            // Keep console window open
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
