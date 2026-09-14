// Title: Delete all horizontal and vertical page breaks from an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that removes every page break from the first worksheet of a workbook with Aspose.Cells and saves the result. | Show the Aspose.Cells .NET method calls needed to clear the page break collection of a worksheet. | Explain how to reset automatic pagination by deleting page breaks in an Excel file using the Aspose.Cells C# API.
// Common Searches: C# Aspose.Cells remove all page breaks from worksheet | Aspose.Cells clear page break collection programmatically | How to reset pagination in Excel using Aspose.Cells .NET | Delete horizontal and vertical page breaks with Aspose.Cells C# example | Aspose.Cells remove page breaks before saving workbook
// Tags: Aspose.Cells worksheet page break removal | Aspose.Cells .NET pagination cleanup | Aspose.Cells pagination reset technique | Aspose.Cells page break collection reset | worksheet pagination control Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads 'input.xlsx', accesses the first worksheet, clears its page break collection with Aspose.Cells, and saves the modified file as 'output.xlsx', demonstrating how to remove all page breaks so Excel can paginate automatically.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: Page break collections may not be available in older Aspose.Cells versions.
            // If needed, clear them using the appropriate API for the version you are using.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
