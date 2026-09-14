// Title: Change a defined name from ReportData to SummaryData in an Excel workbook using Aspose.Cells for .NET and save the file
// AI Prompts: Load an existing .xlsx workbook with Aspose.Cells, locate the defined name "ReportData", set its Text property to "SummaryData", and write the workbook to a new file. | Create C# code that verifies a named range exists, updates its reference string, and persists the changes using Aspose.Cells.
// Common Searches: aspnet rename defined name ReportData to SummaryData using Aspose.Cells | C# update Excel named range Text property with Aspose.Cells | how to change a defined name in an Excel file programmatically Aspose.Cells .NET | Aspose.Cells modify name reference and save workbook
// Tags: Aspose.Cells modify defined name Text property | C# rename Excel defined name Aspose.Cells | update named range reference .NET | save workbook after editing defined name Aspose.Cells | Excel defined name manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads input.xlsx, retrieves the defined name "ReportData", changes its Text property to "SummaryData", and saves the updated workbook as output.xlsx, including checks for file existence and error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Retrieve the defined name "ReportData"
            var name = workbook.Worksheets.Names["ReportData"];
            if (name != null)
            {
                // Change its Text property to "SummaryData"
                name.Text = "SummaryData";
            }
            else
            {
                Console.WriteLine("Defined name \"ReportData\" not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
