// Title: Using Aspose.Cells FindOptions for a case‑sensitive search limited to range E1:E100 in C#
// AI Prompts: Write C# code that uses Aspose.Cells FindOptions with CaseSensitive = true to locate a string only within cells E1 through E100. | Show how to start an Aspose.Cells Find operation at cell E1 and restrict the result to column E rows 1‑100 while enforcing case sensitivity. | Demonstrate verifying that the cell returned by Aspose.Cells Find lies inside the E1:E100 range after a case‑sensitive search.
// Common Searches: Aspose.Cells C# find text case sensitive in column E rows 1 to 100 | How to limit Aspose.Cells Find to a specific range E1:E100 with case sensitivity | C# Aspose.Cells FindOptions CaseSensitive true example for range E1:E100 | Search for exact string in Excel using Aspose.Cells only within E1:E100 | Validate found cell is inside E1:E100 after Aspose.Cells Find
// Tags: Aspose.Cells FindOptions case-sensitive search | limit Aspose.Cells Find to specific range | search column E cells with Aspose.Cells | C# Excel case-sensitive text lookup using Aspose.Cells | validate found cell range Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, configures FindOptions with CaseSensitive = true, starts the search at cell E1, performs a Find limited to column E rows 1‑100, checks that the returned cell falls within the E1:E100 range, outputs the cell address if found, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Define the search text (replace with the actual value you need)
            string searchText = "TargetValue";

            // Configure FindOptions for a case‑sensitive search
            var findOptions = new FindOptions
            {
                CaseSensitive = true
            };

            // Start the search from the first cell of the target range (E1)
            var startCell = worksheet.Cells["E1"];

            // Perform the search; Aspose.Cells provides Find on the Cells collection
            var foundCell = worksheet.Cells.Find(searchText, startCell, findOptions);

            // Verify that the found cell is within the desired range E1:E100
            if (foundCell != null && foundCell.Row <= 99 && foundCell.Column == 4) // Column E = index 4
            {
                Console.WriteLine($"Found at {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("Not found.");
            }

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
