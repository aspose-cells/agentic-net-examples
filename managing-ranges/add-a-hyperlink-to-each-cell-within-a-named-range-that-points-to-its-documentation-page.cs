// Title: Add documentation hyperlinks to each cell of a named range in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through a named range and inserts a hyperlink pointing to a documentation URL built from the cell's address, including a screen tip. | Create a method that receives a workbook path, a named range identifier, and a base URL, then adds a hyperlink to every cell in that range linking to `${baseUrl}{cellAddress}`. | Modify an existing Excel file by programmatically adding per‑cell documentation links to a specified named range and saving the updated file.
// Common Searches: Aspose.Cells C# add hyperlink to each cell in a named range based on cell address | How to generate documentation links for cells in a named range using Aspose.Cells .NET | C# iterate over named range cells and set hyperlink URL with screen tip in Excel | Add per‑cell hyperlinks to a named range in an existing workbook with Aspose.Cells | Aspose.Cells hyperlink URL construction from cell address in a named range
// Tags: Aspose.Cells add hyperlink to named range cells | C# generate cell address URL Aspose.Cells | set hyperlink screentip Aspose.Cells .NET | iterate named range cells Aspose.Cells | Excel workbook modify hyperlinks programmatically

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, retrieves a named range, and adds a hyperlink to each cell that points to a documentation page constructed from the cell's address, then saves the updated workbook.
class AddHyperlinksToNamedRange
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the range to process
            string rangeName = "MyRange";

            // Retrieve the named range
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName(rangeName);
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Worksheet that contains the named range
            Worksheet sheet = namedRange.Worksheet;

            // Base URL for documentation pages
            string docBaseUrl = "https://docs.example.com/";

            // Iterate through each cell in the named range
            for (int row = namedRange.FirstRow; row <= namedRange.FirstRow + namedRange.RowCount - 1; row++)
            {
                for (int col = namedRange.FirstColumn; col <= namedRange.FirstColumn + namedRange.ColumnCount - 1; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Build the hyperlink URL using the cell's address (e.g., A1, B2)
                    string cellAddress = cell.Name; // Returns address like "A1"
                    string hyperlinkUrl = docBaseUrl + cellAddress;

                    // Add the hyperlink (returns the index of the new hyperlink)
                    int hyperlinkIndex = sheet.Hyperlinks.Add(row, col, 1, 1, hyperlinkUrl);
                    Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
                    hyperlink.ScreenTip = "Open documentation for " + cellAddress;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
