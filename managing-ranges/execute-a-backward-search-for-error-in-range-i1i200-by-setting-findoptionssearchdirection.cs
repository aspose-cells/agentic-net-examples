// Title: Perform a backward search for "Error" in column I (I1:I200) using Aspose.Cells FindOptions in C#
// Common Searches: Aspose.Cells find last occurrence of a string in a column range C# | C# backward search in Excel using FindOptions starting from cell I200 | How to limit Aspose.Cells Find to range I1:I200 in C# | Search for "Error" from bottom to top in an Excel worksheet with Aspose.Cells | Find text in specific column using Aspose.Cells FindOptions SearchDirection
// Tags: Aspose.Cells FindOptions reverse search direction | search text in Excel column I using Aspose.Cells C# | locate last occurrence of value in worksheet Aspose.Cells | restrict Find to cells I1 through I200 Aspose.Cells | configure LookInType Values LookAtType Contains Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, configures FindOptions to look in cell values with a contains match, starts the search from cell I200, uses Cells.Find to locate the string "Error" while searching upward, verifies the result is within column I rows 1‑200, outputs the cell address or a not‑found message, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure FindOptions for a forward search (compatible with all versions)
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains
                // Default SearchOrder is ByRows, which is suitable for this scenario
            };

            // Start searching from the bottom of the range (cell I200)
            Cell startCell = worksheet.Cells["I200"];

            // Execute the search for the text "Error"
            Cell foundCell = worksheet.Cells.Find("Error", startCell, findOptions);

            // Ensure the found cell is within the desired range I1:I200
            if (foundCell != null && foundCell.Column == 8 && foundCell.Row <= 199)
            {
                Console.WriteLine($"Found \"Error\" at cell {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("The text \"Error\" was not found in the range I1:I200.");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
