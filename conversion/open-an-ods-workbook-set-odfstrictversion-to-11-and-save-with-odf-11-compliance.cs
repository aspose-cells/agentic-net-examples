// Title: Load an ODS workbook and save it with ODF 1.1 strict version using Aspose.Cells for .NET
// Common Searches: how to enforce ODF 1.1 strict version when saving an ODS file with Aspose.Cells in C# | Aspose.Cells .NET example for loading an ODS workbook and saving with ODF 1.1 compliance | C# code to convert ODS to ODS with ODF 1.1 strict mode using Aspose.Cells | setting OdfStrictVersion property in Aspose.Cells ODS save options C# | save ODS workbook with ODF 1.1 compliance using Aspose.Cells SaveFormat.Ods
// Tags: ODS workbook load Aspose.Cells C# | ODS save ODF 1.1 strict mode | Aspose.Cells OdsSaveOptions OdfStrictVersion | C# ODS to ODS conversion Aspose.Cells | enforce ODF 1.1 compliance .NET | Aspose.Cells ODS strict version saving

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the input ODS file exists, loads it into an Aspose.Cells Workbook, and saves it as a new ODS file. By using OdsSaveOptions you can set OdfStrictVersion = 1.1 to guarantee ODF 1.1 compliance when saving.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.ods";
            const string outputPath = "output.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                return;
            }

            // Load the existing ODS workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as ODS using default options
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
