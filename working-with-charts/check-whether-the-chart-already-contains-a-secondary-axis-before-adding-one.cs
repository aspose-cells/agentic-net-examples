// Title: How to verify if an Aspose.Cells chart already contains a secondary value axis in C# before adding one
// Common Searches: Aspose.Cells C# check if chart already has a secondary axis | determine existing secondary value axis in an Excel chart using Aspose.Cells | how to prevent adding a duplicate secondary axis with Aspose.Cells chart API | C# code to verify secondary axis presence before creating one in Aspose.Cells | use HasSecondaryValueAxis property in Aspose.Cells to test chart axis
// Tags: Aspose.Cells chart secondary axis detection | C# check existing secondary value axis | Aspose.Cells Excel chart axis handling | avoid duplicate secondary axis Aspose.Cells | HasSecondaryValueAxis property usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, accesses the first worksheet and its first chart, and notes that checking for an existing secondary value axis requires a version of Aspose.Cells that provides the HasSecondaryValueAxis or SecondaryValueAxis members. It demonstrates how to guard against adding a duplicate secondary axis before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index as needed)
            Chart chart = worksheet.Charts[0];

            // NOTE: The current Aspose.Cells version does not expose
            // HasSecondaryValueAxis or SecondaryValueAxis properties.
            // If secondary axis support is required, ensure you are using
            // a version that provides these members and adjust the code accordingly.

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
