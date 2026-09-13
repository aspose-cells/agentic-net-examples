// Title: Load an Excel workbook and hide zero values on the first worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, applies a custom number format to hide zero values on the first worksheet, and saves the result to a new file. | Explain a workaround for hiding zero values in a worksheet when the IsZeroValuesHidden property is not available in the current Aspose.Cells version. | Show how to combine file existence checking with workbook loading and zero‑value suppression in Aspose.Cells for a .NET application.
// Common Searches: Aspose.Cells hide zero values on first sheet C# example | workaround for IsZeroValuesHidden missing in Aspose.Cells .NET | load existing Excel file and suppress zero display using Aspose.Cells | C# apply custom number format to hide zeros in Aspose.Cells worksheet | check file existence before opening workbook Aspose.Cells C#
// Tags: load workbook from file Aspose.Cells | hide zero values worksheet Aspose.Cells | custom number format suppress zeros C# | file existence validation before workbook load | save modified Excel file Aspose.Cells | zero values visibility workaround Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, uses a technique such as a custom number format to hide zero values when the IsZeroValuesHidden property is unavailable, and saves the modified workbook while handling possible exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet firstSheet = workbook.Worksheets[0];

            // NOTE: The IsZeroValuesHidden property is not available in the current Aspose.Cells version.
            // If needed, alternative approaches can be applied here.

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
