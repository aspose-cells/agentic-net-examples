using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Source file \"{sourceFile}\" not found.");
            return;
        }

        try
        {
            // LoadOptions can be used for additional settings; external links loading
            // is disabled by default in recent Aspose.Cells versions, so we omit the
            // unavailable property.
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Demonstrate that external links have been ignored (or not loaded)
            Console.WriteLine("External links count after load: " + workbook.Worksheets.ExternalLinks.Count);

            // Save the workbook (optional, just to complete the lifecycle)
            string outputFile = "output.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}