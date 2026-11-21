// Title: How to delete the third chart from an Excel worksheet using Aspose.Cells Charts.RemoveAt in C#
// Common Searches: Aspose.Cells C# remove chart at index 2 from worksheet | Delete specific chart from Excel file with Aspose.Cells Charts.RemoveAt | C# code to remove third chart in a workbook using Aspose.Cells | Remove a chart by position in Aspose.Cells worksheet C# example | Check chart count before removing a chart with Aspose.Cells
// Tags: Aspose.Cells remove chart by index | C# delete third chart worksheet | Charts.RemoveAt Excel workbook | Aspose.Cells chart manipulation C# | Excel worksheet chart removal using Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, verifies that the first worksheet contains at least three charts, removes the chart at zero‑based index 2 with worksheet.Charts.RemoveAt(2), and saves the modified file to a new location.
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
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there are at least three charts before attempting to remove the third one (zero‑based index 2)
            if (worksheet.Charts.Count > 2)
            {
                // Remove the third chart from the worksheet
                worksheet.Charts.RemoveAt(2);
                Console.WriteLine("Third chart removed successfully.");
            }
            else
            {
                Console.WriteLine("The worksheet does not contain a third chart to remove.");
            }

            // Save the workbook with the chart removed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
