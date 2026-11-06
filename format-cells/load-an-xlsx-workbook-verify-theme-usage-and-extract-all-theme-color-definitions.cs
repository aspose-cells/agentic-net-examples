// Title: Load an XLSX workbook in C# with Aspose.Cells and read the workbook's Theme property
// Common Searches: how to read the Theme property of an Excel workbook using Aspose.Cells in C# | C# Aspose.Cells check if an XLSX file contains a theme | retrieve workbook theme name with Aspose.Cells .NET | determine whether an Excel file has a theme applied using Aspose.Cells
// Tags: Aspose.Cells read workbook theme | C# Aspose.Cells detect Excel theme | Aspose.Cells Theme property usage | load XLSX workbook Aspose.Cells .NET | verify theme presence Excel file

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example verifies that an input.xlsx file exists, loads it into an Aspose.Cells Workbook, checks the Workbook.Theme property, and outputs the theme name if available or a message indicating that no theme information is present, while handling any runtime exceptions.
class ThemeExtractor
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to retrieve theme information (available in newer versions)
            // If the Theme property returns a non‑empty string, display it.
            // Detailed theme color extraction requires ThemeColorScheme which may not be present in older versions.
            if (!string.IsNullOrEmpty(workbook.Theme))
            {
                Console.WriteLine($"Workbook theme: {workbook.Theme}");
            }
            else
            {
                Console.WriteLine("The workbook does not contain a theme or theme information is unavailable in this version of Aspose.Cells.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
