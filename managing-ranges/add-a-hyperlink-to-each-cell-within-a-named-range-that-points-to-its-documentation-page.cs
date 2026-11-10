// Title: Add documentation hyperlinks to every cell in a named range with Aspose.Cells for .NET (C#)
// Common Searches: how to add a hyperlink to each cell in a named range using Aspose.Cells C# | Aspose.Cells C# add cell‑specific documentation links based on cell address | programmatically create hyperlinks for all cells of a named range in an Excel file | using GetRangeByName to iterate cells and add hyperlinks with Aspose.Cells | C# Aspose.Cells example for adding Hyperlinks.Add entries to a range of cells
// Tags: cell hyperlink creation Aspose.Cells | named range iteration Aspose.Cells | documentation URL generation C# | Excel hyperlink addition .NET | Hyperlinks.Add API Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, retrieves the named range "MyRange", iterates through each cell in that range, builds a documentation URL from the cell address, adds a hyperlink to the cell using the worksheet's Hyperlinks.Add method, sets the ScreenTip and display text, and saves the updated file.
class AddHyperlinksToNamedRange
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range definition
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Resolve the named range to an actual cell range
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName("MyRange");
            if (range == null)
            {
                Console.WriteLine("Unable to resolve the range for 'MyRange'.");
                return;
            }

            // Base URL for documentation pages
            string baseUrl = "https://docs.example.com/";

            // Iterate through each cell in the range and add a hyperlink
            for (int row = range.FirstRow; row <= range.FirstRow + range.RowCount - 1; row++)
            {
                for (int col = range.FirstColumn; col <= range.FirstColumn + range.ColumnCount - 1; col++)
                {
                    try
                    {
                        // Build the URL using the cell address (e.g., "A1")
                        string cellAddress = CellsHelper.CellIndexToName(row, col);
                        string docUrl = baseUrl + cellAddress;

                        // Add a hyperlink to the cell via the worksheet's Hyperlinks collection
                        int hyperlinkIndex = sheet.Hyperlinks.Add(row, col, 1, 1, docUrl);
                        Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
                        hyperlink.ScreenTip = docUrl;
                        hyperlink.TextToDisplay = docUrl;
                    }
                    catch (Exception exCell)
                    {
                        Console.WriteLine($"Failed to add hyperlink to cell [{row}, {col}]: {exCell.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
