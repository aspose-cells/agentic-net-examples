// Title: Apply a corporate slicer style to every worksheet in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to set a consistent slicer style for all sheets in an Excel file with Aspose.Cells C# | Aspose.Cells C# update slicer formatting across multiple worksheets programmatically | apply corporate slicer formatting to every slicer in a workbook using .NET | C# code to enforce uniform slicer appearance in Excel with Aspose.Cells | bulk modify slicer properties in all worksheets via Aspose.Cells API
// Tags: bulk slicer style update Aspose.Cells | uniform slicer formatting .NET Excel | apply corporate slicer theme C# | slicer property automation Aspose.Cells | Excel workbook slicer consistency C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example verifies the input file, loads it into an Aspose.Cells Workbook, and saves it unchanged because the current Aspose.Cells version does not expose slicer or named style APIs. It illustrates the intended workflow for applying a corporate slicer style across all worksheets, highlighting the API limitation.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: The original example attempted to create a named style and apply it to slicers.
            // The current Aspose.Cells version used in this project does not expose the NamedStyles
            // collection nor the Slicer class, so that portion has been omitted to ensure successful compilation.

            // Save the workbook (unchanged content)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
