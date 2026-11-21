// Title: How to prevent users from inserting new rows into an Excel table (list object) with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# disable adding rows to a list object | prevent row insertion in Excel table with Aspose.Cells .NET | protect worksheet to stop users adding rows in Aspose.Cells | fixed-size Excel table programmatically using Aspose.Cells | Aspose.Cells prevent users from expanding a table
// Tags: sheet-level protection block row addition Aspose.Cells | list object insertion restriction using Aspose.Cells C# | non-expandable Excel table implementation Aspose.Cells | apply full protection type All Aspose.Cells | prevent table growth programmatically Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a workbook, checks for list objects on the first worksheet, and applies full worksheet protection (ProtectionType.All) to disable row insertion, then saves the workbook.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // If the worksheet contains any list objects (tables), protect it
                if (worksheet.ListObjects.Count > 0)
                {
                    // Protect the worksheet; this disables row insertion among other actions
                    worksheet.Protect(ProtectionType.All);

                    // Additional permissions can be adjusted here if needed, e.g.:
                    // worksheet.Protection.AllowEditObject = true;
                    // worksheet.Protection.AllowEditScenario = true;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
