// Title: Detect Header‑Only Worksheets and Flag Them with Aspose.Cells for .NET (C#)
// Description: A C# console program that loads an Excel workbook using Aspose.Cells, scans each worksheet, and flags any sheet that contains only a header row (or is completely empty). It adds a string custom property "ReviewFlag" set to "true", writes a log message to the console, and saves the updated file.
// Keywords: Aspose.Cells header only worksheet | detect empty rows Aspose.Cells | C# worksheet validation | custom property ReviewFlag | MaxDataRow usage | flag Excel sheet for review | Excel quality control automation | Aspose.Cells workbook processing | detect template sheets | Excel data integrity check
// Common Searches: How to find worksheets with only header rows using Aspose.Cells | Add a custom property to flag empty Excel sheets in C# | Aspose.Cells detect and log header‑only worksheets | Mark template worksheets for review with Aspose.Cells | C# code to check MaxDataRow in Excel workbook
// Developer Intent: Identify worksheets that contain only a header row, mark them with a custom property, and log the occurrence.
// Use Cases: Automatically flag template or placeholder sheets that have not been populated before publishing. | Generate a pre‑publish report of empty or header‑only worksheets for data‑quality audits. | Enable downstream processes to skip or specially handle sheets flagged with "ReviewFlag".
// AI Prompts: Create a method that returns a list of worksheet names that are header‑only using Aspose.Cells. | Extend the example to change the tab color of flagged worksheets to red. | Write code to remove the "ReviewFlag" custom property after a sheet has been reviewed.

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetHeaderChecker
{
    // A C# console program that loads an Excel workbook using Aspose.Cells, scans each worksheet, and flags any sheet that contains only a header row (or is completely empty). It adds a string custom property "ReviewFlag" set to "true", writes a log message to the console, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the last row that contains any data (-1 if the sheet is completely empty)
                    int lastDataRow = sheet.Cells.MaxDataRow;

                    // If there is no data or only the first row has data, treat it as header‑only
                    bool isHeaderOnly = lastDataRow <= 0;

                    if (isHeaderOnly)
                    {
                        // Flag the worksheet for review using a custom property (value must be a string)
                        sheet.CustomProperties.Add("ReviewFlag", "true");

                        // Log relevant details
                        Console.WriteLine($"Worksheet '{sheet.Name}' contains only header rows and has been flagged for review.");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
