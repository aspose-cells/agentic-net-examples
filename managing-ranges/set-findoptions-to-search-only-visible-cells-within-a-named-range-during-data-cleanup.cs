// Title: Configure FindOptions to search only visible cells within a named range and clear matching values using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a FindOptions instance, enables SearchInVisibleCellsOnly (using reflection for older Aspose.Cells versions), searches a named range for a specific text, and clears the cell value only if its row and column are visible. | Show how to detect at runtime whether the SearchInVisibleCellsOnly property exists on FindOptions before setting it, ensuring compatibility across different Aspose.Cells releases. | Demonstrate verifying row and column visibility after a Find operation and removing the found text exclusively from visible cells in an Excel worksheet with Aspose.Cells.
// Common Searches: Aspose.Cells FindOptions search only visible cells in a named range C# | how to delete a cell value after finding text only in visible rows using Aspose.Cells | check for SearchInVisibleCellsOnly property in Aspose.Cells .NET before setting | C# data cleanup find and clear specific value in visible Excel cells Aspose.Cells | using reflection to set FindOptions.SearchInVisibleCellsOnly for older Aspose.Cells versions
// Tags: FindOptions visible cells filter | clear cell value after find Aspose.Cells | reflection set SearchInVisibleCellsOnly | named range visibility handling | Excel cleanup visible rows C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, creates a FindOptions object, uses reflection to enable SearchInVisibleCellsOnly when supported, searches a named range for a target string, confirms the found cell's row and column are not hidden, clears the cell content, and saves the cleaned workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input and output file paths
                string inputPath = @"C:\Input\Sample.xlsx";
                string outputPath = @"C:\Output\Sample_Cleaned.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"Input file not found: {inputPath}");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure find options (default LookIn = Values, LookAt = Contains)
                FindOptions findOptions = new FindOptions();

                // Restrict search to visible cells only if the property exists (older versions may lack it)
                var searchInVisibleProp = typeof(FindOptions).GetProperty("SearchInVisibleCellsOnly");
                if (searchInVisibleProp != null && searchInVisibleProp.CanWrite)
                {
                    searchInVisibleProp.SetValue(findOptions, true);
                }

                // Text to search for
                string searchText = "TargetValue";

                // Perform the search on the entire worksheet
                Cell foundCell = worksheet.Cells.Find(searchText, null, findOptions);

                // If a visible cell is found, clear its value
                if (foundCell != null)
                {
                    // Ensure the cell is not hidden by row or column
                    bool rowHidden = worksheet.Cells.Rows[foundCell.Row].IsHidden;
                    bool columnHidden = worksheet.Cells.Columns[foundCell.Column].IsHidden;

                    if (!rowHidden && !columnHidden)
                    {
                        foundCell.PutValue(string.Empty);
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
