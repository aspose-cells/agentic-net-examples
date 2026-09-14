// Title: Save an Aspose.Cells workbook as CSV with TxtSaveOptions that trim leading blank rows and columns (C#)
// AI Prompts: Write C# code that creates a workbook, adds data starting at cell C3, configures TxtSaveOptions to remove leading empty rows and columns, sets a comma delimiter and UTF‑8 encoding, and saves the result as a CSV file. | Show how to use Aspose.Cells TxtSaveOptions in C# to export a worksheet to CSV while automatically trimming blank rows/columns and specifying a custom separator.
// Common Searches: asp.net aspose.cells csv trim leading blank rows and columns | c# txtsaveoptions trimleadingblankrowandcolumn example | export excel worksheet to csv without empty rows using Aspose.Cells | set custom csv delimiter and utf-8 encoding in Aspose.Cells C# | how to remove empty rows and columns when saving to CSV with Aspose.Cells
// Tags: Aspose.Cells TxtSaveOptions CSV trimming | TrimLeadingBlankRowAndColumn export | C# Aspose.Cells set CSV separator | UTF-8 encoding Aspose.Cells CSV output | remove empty rows Aspose.Cells conversion

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvTrimExample
{
    // Creates a workbook, places data at C3, configures TxtSaveOptions to trim leading blank rows and columns, sets a comma separator and UTF‑8 encoding, and saves the workbook as a trimmed CSV file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add data with leading blank rows and columns
            // Row 0 and 1 are blank, column 0 and 1 are blank
            cells["C3"].PutValue("First");
            cells["D4"].PutValue("Second");
            cells["E5"].PutValue("Third");

            // Configure TxtSaveOptions to trim leading blank rows and columns
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                // Ensure leading blanks are removed (default is true, set explicitly for clarity)
                TrimLeadingBlankRowAndColumn = true,
                // Use comma as the CSV separator
                Separator = ',',
                // Optional: set encoding if needed
                Encoding = Encoding.UTF8
            };

            // Save the workbook as CSV with the specified options
            workbook.Save("TrimmedOutput.csv", saveOptions);

            Console.WriteLine("Workbook saved as CSV with leading blanks trimmed.");
        }
    }
}
