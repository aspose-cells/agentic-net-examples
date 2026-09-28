// Title: Configure Aspose.Cells FindOptions in C# to skip hidden rows when searching the K1:K500 range
// AI Prompts: Write C# code that sets FindOptions.IncludeHiddenRows = false and searches for a text value only within cells K1:K500 using Aspose.Cells. | Show how to create a Range for K1:K500, apply FindOptions to ignore hidden rows, and retrieve the first matching cell. | Demonstrate validating that the cell returned by Worksheet.Cells.Find lies inside the specified range after hidden‑row filtering.
// Common Searches: Aspose.Cells C# find text in column K ignoring hidden rows | How to use FindOptions to exclude hidden rows in Aspose.Cells | Search range K1:K500 with Aspose.Cells Find method C# | C# Aspose.Cells find with hidden row filter example | Limit Aspose.Cells Find to specific column and skip hidden rows
// Tags: Aspose.Cells FindOptions ignore hidden rows | search specific range K1:K500 Aspose.Cells | C# Excel hidden rows filtering Aspose.Cells | cell find operation within column K Aspose.Cells | range-based find with hidden row exclusion C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads a workbook, creates a range covering cells K1:K500, configures FindOptions to exclude hidden rows, searches for a specified text using Worksheet.Cells.Find, verifies the found cell is inside the defined range, outputs the result, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the search range K1:K500 (zero‑based indices: row 0, column 10, 500 rows, 1 column)
            AsposeRange searchRange = sheet.Cells.CreateRange(0, 10, 500, 1);

            // Configure FindOptions (default options are sufficient for a simple text search)
            FindOptions findOptions = new FindOptions();

            // Perform the search using the worksheet's Cells collection
            // The Find method searches the whole sheet, so we later verify the cell is inside our range
            Cell foundCell = sheet.Cells.Find("searchText", null, findOptions);

            // Verify that the found cell lies within the defined range
            if (foundCell != null &&
                foundCell.Row >= searchRange.FirstRow && foundCell.Row < searchRange.FirstRow + searchRange.RowCount &&
                foundCell.Column >= searchRange.FirstColumn && foundCell.Column < searchRange.FirstColumn + searchRange.ColumnCount)
            {
                Console.WriteLine($"Found at {foundCell.Name} with value: {foundCell.StringValue}");
            }
            else
            {
                Console.WriteLine("Value not found in the specified range.");
            }

            // Save the workbook if any modifications were made (optional)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
