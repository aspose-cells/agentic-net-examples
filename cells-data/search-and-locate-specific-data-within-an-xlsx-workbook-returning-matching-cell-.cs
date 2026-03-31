using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSearchDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX workbook
            string workbookPath = "input.xlsx";

            // The value to search for (can be string, int, double, DateTime, bool)
            string searchValue = "Apple";

            // Load the workbook (lifecycle rule: use constructor)
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (you can modify to target a specific sheet)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure find options: search in cell values, look for containing text
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                SearchBackward = false,
                SeachOrderByRows = true
            };

            // List to hold results
            List<(string Address, object Value)> matches = new List<(string, object)>();

            // Perform iterative search using the Find method
            Cell previousCell = null;
            while (true)
            {
                // Find the next cell containing the search value
                Cell foundCell = cells.Find(searchValue, previousCell, findOptions);

                // If no more matches, exit loop
                if (foundCell == null)
                    break;

                // Store the cell address (e.g., "A2") and its raw value
                matches.Add((foundCell.Name, foundCell.Value));

                // Set previousCell to the current found cell for the next iteration
                previousCell = foundCell;
            }

            // Output the results
            Console.WriteLine($"Search results for \"{searchValue}\":");
            foreach (var match in matches)
            {
                Console.WriteLine($"Cell {match.Address} = {match.Value}");
            }

            // Optionally, save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}