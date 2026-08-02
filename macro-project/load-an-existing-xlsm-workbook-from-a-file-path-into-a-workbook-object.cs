// Title: Load an XLSM workbook from a file path with Aspose.Cells for .NET
// Description: C# example that checks for a macro‑enabled .xlsm file, loads it into an Aspose.Cells Workbook when present, or creates a new workbook with a default sheet when missing. The code then prints the first worksheet name, total sheet count, and handles any exceptions.
// Keywords: Aspose.Cells C# | .NET load xlsm | macro enabled workbook | open workbook from path | file existence check | create new workbook Aspose.Cells | read worksheet name | count worksheets | exception handling
// Common Searches: Aspose.Cells open .xlsm file C# | load macro enabled workbook from path | check if Excel file exists before loading Aspose.Cells | create workbook when file not found Aspose.Cells | get first sheet name Aspose.Cells
// Developer Intent: Open a macro‑enabled Excel file from disk, with a safe fallback to a new workbook if the file is absent.
// Use Cases: Read the name of the first worksheet in an existing .xlsm file. | Display the total number of worksheets after loading a workbook. | Automatically generate a new workbook with a default sheet when the target file does not exist. | Capture and log errors that occur during workbook loading.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an .xlsm file from a given path, creates a new workbook if the file is missing, and prints the first sheet name and sheet count. | Show how to combine File.Exists with Aspose.Cells Workbook constructor and exception handling for macro‑enabled Excel files. | Write a method that returns a Workbook object after loading an existing .xlsm or initializing a new one with a default worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // C# example that checks for a macro‑enabled .xlsm file, loads it into an Aspose.Cells Workbook when present, or creates a new workbook with a default sheet when missing. The code then prints the first worksheet name, total sheet count, and handles any exceptions.
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSM workbook (adjust as needed)
            string workbookPath = @"C:\Path\To\YourWorkbook.xlsm";

            Workbook workbook = null;

            try
            {
                if (File.Exists(workbookPath))
                {
                    // Load the existing workbook
                    workbook = new Workbook(workbookPath);
                }
                else
                {
                    // File not found – create a new workbook with a default sheet
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                    Console.WriteLine($"File not found: '{workbookPath}'. A new workbook has been created.");
                }

                // Access the first worksheet and display its name
                Worksheet firstSheet = workbook.Worksheets[0];
                Console.WriteLine("First worksheet name: " + firstSheet.Name);

                // Display the total number of worksheets
                Console.WriteLine("Total worksheets: " + workbook.Worksheets.Count);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook: " + ex.Message);
            }
        }
    }
}
