// Title: Perform a case‑insensitive substring search for "FY" in an Excel worksheet range using Aspose.Cells FindOptions (C#)
// Common Searches: how to use Aspose.Cells FindOptions for case‑insensitive text search in C# | search for a substring in a specific Excel range with Aspose.Cells C# | find cells containing FY regardless of case using Aspose.Cells Find method | Aspose.Cells C# example of FindOptions LookAtType.Contains with MaxDisplayRange | retrieve first cell that contains abbreviation FY in an Excel workbook using Aspose.Cells
// Tags: case-insensitive FindOptions search Aspose.Cells C# | Find method substring lookup Aspose.Cells | search within MaxDisplayRange Aspose.Cells | LookAtType.Contains usage Aspose.Cells | Excel workbook cell search Aspose.Cells C#

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, populates cells with different case variations of "FY", defines the used range via MaxDisplayRange, configures FindOptions to search cell values using LookAtType.Contains while keeping MatchCase false, and calls Cells.Find to locate the first cell containing the substring "FY" regardless of case, then outputs the cell address and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data containing different case variations of "FY"
            sheet.Cells["A1"].PutValue("FY2021");
            sheet.Cells["A2"].PutValue("fy2022");
            sheet.Cells["A3"].PutValue("Q1");
            sheet.Cells["B1"].PutValue("Report");
            sheet.Cells["B2"].PutValue("fy Summary");
            sheet.Cells["B3"].PutValue("FY");

            // Define the range to search (entire used range)
            AsposeRange searchRange = sheet.Cells.MaxDisplayRange;

            // Configure FindOptions for a case‑insensitive search
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,      // Search cell values
                LookAtType = LookAtType.Contains      // Look for substring
                // MatchCase defaults to false (case‑insensitive)
            };

            // Perform the search for the abbreviation "FY"
            Cell foundCell = null;
            try
            {
                // Cells.Find overload requires a start cell; passing null starts from the first cell
                foundCell = sheet.Cells.Find("FY", null, findOptions);
            }
            catch (Exception findEx)
            {
                Console.WriteLine($"Find operation error: {findEx.Message}");
            }

            if (foundCell != null)
            {
                // Output the address of the first match
                Console.WriteLine($"Found at: {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("No match found.");
            }

            // Save the workbook (optional)
            try
            {
                workbook.Save("FindFY.xlsx");
                Console.WriteLine("Workbook saved as FindFY.xlsx");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
