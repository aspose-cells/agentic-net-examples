// Title: C# – Load an Excel workbook, update a named range, and save as a new file using Aspose.Cells
// Description: This example checks for the presence of Report.xlsx, opens it with Aspose.Cells for .NET, retrieves the named range "MyRange", replaces every cell value in that range with "Updated", and writes the result to ReportUpdated.xlsx. It includes basic file‑existence and range‑validation handling.
// Keywords: Aspose.Cells C# | update named range | modify Excel cells .NET | load workbook Aspose | save workbook as new file | named range iteration | Excel file existence check | Report.xlsx | ReportUpdated.xlsx | error handling Aspose.Cells
// Common Searches: Aspose.Cells change values in a named range C# | How to save a modified workbook with a different name using Aspose.Cells | Check if a named range exists before updating Aspose.Cells .NET | Iterate through cells of a named range in C# | Replace all cells in a named range with a constant string Aspose
// Developer Intent: Replace every cell in a specific named range of an existing Excel file and create a new workbook with the changes applied.
// Use Cases: Refresh a report section by overwriting a predefined named range with a new value. | Create versioned copies of a workbook after bulk updates while keeping the original intact. | Validate named‑range presence to prevent runtime errors during automated data processing.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, updates all cells in a given named range to a supplied string, and saves the workbook under a new filename. | Explain best practices for handling missing files or undefined named ranges when using Aspose.Cells in .NET. | Show how to log the address of each cell modified inside a named range with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This example checks for the presence of Report.xlsx, opens it with Aspose.Cells for .NET, retrieves the named range "MyRange", replaces every cell value in that range with "Updated", and writes the result to ReportUpdated.xlsx. It includes basic file‑existence and range‑validation handling.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "Report.xlsx";
                const string outputPath = "ReportUpdated.xlsx";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range (replace "MyRange" with the actual name in the workbook)
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' not found.");
                    return;
                }

                // Get the actual cell range that the name refers to
                Aspose.Cells.Range range = namedRange.GetRange();

                // Modify each cell in the range – set a new string value
                foreach (Cell cell in range)
                {
                    cell.PutValue("Updated");
                }

                // Save the modified workbook to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
