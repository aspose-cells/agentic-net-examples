// Title: C# Aspose.Cells Example – Load Report.xlsx, Update a Named Range, and Save as ReportUpdated.xlsx
// Description: A concise C# sample that checks for Report.xlsx, opens it with Aspose.Cells, retrieves the named range "MyRange", overwrites every cell in that range with the text "Updated", and saves the result as ReportUpdated.xlsx. Includes robust error handling for missing files or ranges.
// Keywords: Aspose.Cells C# load workbook | modify named range Aspose.Cells | update cells in named range .NET | save workbook as new file Aspose | Excel named range example C# | bulk cell value change Aspose.Cells | error handling Aspose.Cells
// Common Searches: Aspose.Cells update named range C# | load Excel file and change named range with Aspose | C# code to edit a named range and save a new workbook | Aspose.Cells example for modifying cells in a named range | how to replace values in a named range using Aspose.Cells
// Developer Intent: Replace the contents of a specific named range in an existing Excel workbook and write the modified workbook to a new file.
// Use Cases: Refresh a template by overwriting a predefined named range with a constant value before distribution. | Generate a customized report version by bulk‑updating a named range with new data. | Validate the existence of a named range, modify its cells, and create a versioned copy for audit trails.
// AI Prompts: Create C# code using Aspose.Cells to open 'Report.xlsx', set every cell in the named range 'MyRange' to "Updated", and save as 'ReportUpdated.xlsx'. | Show how to add comprehensive error handling for missing source files or undefined named ranges when updating an Excel range with Aspose.Cells. | Provide an Aspose.Cells snippet that iterates over a named range and applies custom formatting (e.g., background color) to each cell.

using System;
using System.IO;
using Aspose.Cells;

namespace ReportUpdateExample
{
    // A concise C# sample that checks for Report.xlsx, opens it with Aspose.Cells, retrieves the named range "MyRange", overwrites every cell in that range with the text "Updated", and saves the result as ReportUpdated.xlsx. Includes robust error handling for missing files or ranges.
    class Program
    {
        static void Main()
        {
            const string inputPath = "Report.xlsx";
            const string outputPath = "ReportUpdated.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range (replace "MyRange" with the actual name in the workbook)
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range \"MyRange\" not found in the workbook.");
                    return;
                }

                // Get the actual cell range that the name refers to
                var range = namedRange.GetRange();

                // Example modification: set every cell in the range to the text "Updated"
                foreach (Cell cell in range)
                {
                    cell.PutValue("Updated");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
