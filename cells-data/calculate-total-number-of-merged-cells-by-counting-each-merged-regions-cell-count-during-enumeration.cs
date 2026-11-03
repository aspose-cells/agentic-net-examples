// Title: How to compute the total number of merged cells in every worksheet of an XLSX file using Aspose.Cells for .NET (C#)
// Common Searches: C# code to sum all merged cells in an XLSX workbook with Aspose.Cells | Aspose.Cells GetMergedAreas example for counting merged cells across worksheets | How to handle password‑protected Excel files while counting merged cells using Aspose.Cells | Calculate merged region cell total in a multi‑sheet workbook with Aspose.Cells .NET | Retrieve total merged cell count from an Excel file using Aspose.Cells API
// Tags: Aspose.Cells GetMergedAreas merged cell count | count merged cells across worksheets Aspose.Cells | load password protected XLSX Aspose.Cells .NET | enumerate merged regions Excel C# Aspose.Cells | calculate merged region cell total Aspose.Cells | Aspose.Cells workbook merged area aggregation

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an XLSX workbook (including optional password handling), iterates through each worksheet, uses Cells.GetMergedAreas() to enumerate every merged region, computes each region's row and column span, aggregates the cell count of all merged areas, and prints the total number of merged cells in the entire workbook.
class Program
{
    static void Main()
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException)
        {
            // Assume the workbook is password‑protected; prompt the user for a password
            Console.WriteLine("The workbook is password protected.");
            Console.Write("Enter password: ");
            string password = Console.ReadLine();

            try
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Failed to load workbook with provided password: {innerEx.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        if (workbook == null)
        {
            Console.WriteLine("Workbook could not be loaded.");
            return;
        }

        try
        {
            int totalMergedCells = 0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Enumerate all merged cell areas using the recommended API
                foreach (CellArea mergedArea in sheet.Cells.GetMergedAreas())
                {
                    // Calculate the number of rows and columns in the merged region
                    int rowCount = mergedArea.EndRow - mergedArea.StartRow + 1;
                    int columnCount = mergedArea.EndColumn - mergedArea.StartColumn + 1;

                    // Add the cell count of this merged region to the total
                    totalMergedCells += rowCount * columnCount;
                }
            }

            // Output the total number of merged cells
            Console.WriteLine($"Total merged cells: {totalMergedCells}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
