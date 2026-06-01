using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In some Aspose.Cells versions the LoadWarnings collection is not available.
            // If needed, warnings can be retrieved via other mechanisms provided by the library.

            Console.WriteLine("Workbook loaded successfully.");

            // (Optional) Save the workbook if further processing is needed
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}