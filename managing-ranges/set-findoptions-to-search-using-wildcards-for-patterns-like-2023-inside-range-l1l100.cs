// Title: C# Aspose.Cells FindOptions with Excel‑style wildcards to locate "*2023*" in column L (L1:L100)
// Description: Demonstrates how to configure FindOptions for Excel‑style wildcard matching, set a CellArea covering L1:L100, and use Cells.Find to return the first cell whose value contains the pattern "*2023*" in a .NET workbook.
// Keywords: Aspose.Cells | FindOptions | wildcard search | C# .NET | Excel wildcards | search range L1:L100 | CellArea | LookAtType.Contains | SetRange | Cells.Find
// Common Searches: Aspose.Cells FindOptions wildcard example | search column L for "*2023*" using Aspose.Cells | set range for Find method Aspose.Cells C# | enable Excel‑style wildcards in Aspose.Cells Find | find cells containing year 2023 in .NET workbook
// Developer Intent: Retrieve the first cell in column L (rows 1‑100) whose text includes the substring 2023 by applying an Excel‑style wildcard pattern with FindOptions.
// Use Cases: Filter rows that belong to a specific fiscal year before processing data. | Validate that a worksheet contains entries for a given year within a designated column. | Extract or highlight cells matching a year pattern prior to exporting or reporting.
// AI Prompts: Generate C# code using Aspose.Cells FindOptions to locate cells containing "*2024*" in range B2:B500 with case‑insensitive wildcard matching. | Explain how to configure FindOptions for a multi‑column, case‑insensitive wildcard search that returns all matching cells in Aspose.Cells. | Show how to combine FindOptions with conditional formatting to highlight every cell in column L that matches the pattern "*2023*".

using System;
using Aspose.Cells;

// Demonstrates how to configure FindOptions for Excel‑style wildcard matching, set a CellArea covering L1:L100, and use Cells.Find to return the first cell whose value contains the pattern "*2023*" in a .NET workbook.
class FindWithWildcards
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample data in column L (index 11) for demonstration
        cells["L1"].PutValue("Report_2022");
        cells["L2"].PutValue("Summary_2023_Q1");
        cells["L3"].PutValue("Data_2023_Final");
        cells["L4"].PutValue("Archive_2021");

        // Create FindOptions instance
        FindOptions findOptions = new FindOptions();

        // Define the search range L1:L100
        CellArea searchArea = new CellArea
        {
            StartRow = 0,          // Row 1 (0‑based)
            StartColumn = 11,      // Column L (0‑based)
            EndRow = 99,           // Row 100
            EndColumn = 11         // Column L
        };
        findOptions.SetRange(searchArea);

        // Search in cell values and use Excel‑style wildcards
        findOptions.LookInType = LookInType.Values;      // Search in values
        findOptions.LookAtType = LookAtType.Contains;    // Enables wildcard handling
        findOptions.RegexKey = false;                    // Use Excel wildcards, not regex

        // Perform the find operation with the wildcard pattern "*2023*"
        Cell foundCell = cells.Find("*2023*", null, findOptions);

        // Output the result
        if (foundCell != null)
        {
            Console.WriteLine($"Found cell: {foundCell.Name} with value \"{foundCell.StringValue}\"");
        }
        else
        {
            Console.WriteLine("No cell matching the pattern was found.");
        }

        // (Optional) Save the workbook if needed
        // workbook.Save("FindWithWildcards.xlsx");
    }
}
