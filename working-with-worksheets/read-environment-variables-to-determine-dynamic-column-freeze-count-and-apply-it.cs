// Title: Use an environment variable to set a dynamic column freeze count with Aspose.Cells FreezePanes in C#
// Common Searches: C# Aspose.Cells freeze first N columns using an environment variable | How to read FREEZE_COLUMN_COUNT and apply FreezePanes in .NET | Set dynamic column freeze panes in an Excel workbook with Aspose.Cells | Programmatically freeze columns based on env variable in Aspose.Cells for C# | Aspose.Cells FreezePanes overload example reading column count from environment
// Tags: environment variable column freeze Aspose.Cells | FreezePanes overload with column count C# | dynamic column freeze Excel workbook | read environment variable Aspose.Cells worksheet | programmatic column freeze Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, reads the FREEZE_COLUMN_COUNT environment variable, parses it to an integer, and uses worksheet.FreezePanes(0, count, 0, count) to freeze that many columns before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; if not, create a new workbook as a fallback.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Read the environment variable that specifies how many columns to freeze.
            string envValue = Environment.GetEnvironmentVariable("FREEZE_COLUMN_COUNT");
            int freezeColumnCount = 0;

            if (!string.IsNullOrEmpty(envValue) && int.TryParse(envValue, out int parsedValue))
            {
                freezeColumnCount = parsedValue;
            }

            // Apply freeze panes: freeze rows = 0, columns = freezeColumnCount.
            // Use the overload that includes total rows/columns to freeze.
            worksheet.FreezePanes(0, freezeColumnCount, 0, freezeColumnCount);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
