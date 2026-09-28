// Title: How to set WinForms ListBox selected indices from Excel column A using Aspose.Cells in C#
// AI Prompts: Write a C# method that opens an Excel workbook with Aspose.Cells, reads every non‑empty cell in column A, and returns a List<int> of indices where the cell values match strings in a supplied IList<string>. | Generate code that takes the list of matching indices and assigns them to the SelectedIndices collection of a WinForms ListBox.
// Common Searches: c# aspose.cells read column a and get matching list indices | winforms listbox select items based on excel column values | map excel column data to listbox selectedindices using aspose.cells | retrieve row numbers from column a that match a list of strings in c#
// Tags: Aspose.Cells column A extraction in C# | ListBox SelectedIndices assignment from Excel data | C# workbook row matching with string list | WinForms ListBox population using spreadsheet values | Excel-to-UI index mapping with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook via Aspose.Cells, iterates through used rows in column A, trims each cell value, and checks for its presence in a provided list of strings. Matching item indices are collected and returned, enabling those indices to be applied to a WinForms ListBox's SelectedIndices collection. Includes file‑existence validation and basic exception handling.
public static class ExcelListBoxHelper
{
    /// <param name="excelFilePath">Full path to the Excel workbook.</param>
    /// <param name="items">A list of items (e.g., ListBox items) to match against the Excel values.</param>
    /// <returns>A list of indices that correspond to items found in column A.</returns>
    public static List<int> GetMatchingIndicesFromColumnA(string excelFilePath, IList<string> items)
    {
        var matchingIndices = new List<int>();

        try
        {
            // Verify that the Excel file exists to avoid FileNotFoundException.
            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException($"The file '{excelFilePath}' was not found.");

            // Load the workbook using Aspose.Cells.
            var workbook = new Workbook(excelFilePath);

            // Access the first worksheet (index 0).
            var worksheet = workbook.Worksheets[0];

            // Get the cells collection for easier access.
            var cells = worksheet.Cells;

            // Determine the last row that contains data in column A (zero‑based index).
            int lastRow = cells.MaxDataRow;

            // Iterate through each used row in column A.
            for (int row = 0; row <= lastRow; row++)
            {
                // Retrieve the cell value as a trimmed string (may be null).
                string? cellValue = cells[row, 0].StringValue?.Trim();

                // Skip empty cells.
                if (string.IsNullOrEmpty(cellValue))
                    continue;

                // Find the index of the item in the provided list that matches the cell value.
                int itemIndex = items.IndexOf(cellValue);

                // If the item exists, add its index to the result list.
                if (itemIndex >= 0 && !matchingIndices.Contains(itemIndex))
                {
                    matchingIndices.Add(itemIndex);
                }
            }
        }
        catch (Exception ex)
        {
            // Log or handle the exception as needed.
            Console.Error.WriteLine($"Error processing Excel file: {ex.Message}");
            // Optionally rethrow or return an empty list.
        }

        return matchingIndices;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Example usage:
            // Provide the path to the Excel file and a sample list of items.
            string excelPath = "SampleData.xlsx";

            // Ensure the file exists before proceeding.
            if (!File.Exists(excelPath))
            {
                Console.Error.WriteLine($"Excel file not found: {excelPath}");
                return;
            }

            var items = new List<string> { "Item1", "Item2", "Item3", "Item4" };

            List<int> matchingIndices = ExcelListBoxHelper.GetMatchingIndicesFromColumnA(excelPath, items);

            Console.WriteLine("Matching indices:");
            foreach (int index in matchingIndices)
            {
                Console.WriteLine(index);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
