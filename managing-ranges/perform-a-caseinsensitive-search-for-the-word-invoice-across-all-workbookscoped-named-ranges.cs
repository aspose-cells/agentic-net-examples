// Title: C# – Case‑Insensitive Search for "invoice" in All Workbook‑Scoped Named Ranges (Aspose.Cells)
// Description: Creates a workbook, defines global named ranges, configures FindOptions for a case‑insensitive "contains" search, filters workbook‑scoped names, iterates each range, finds every "invoice" occurrence, writes results to the console, and saves the file.
// Keywords: Aspose.Cells | C# | FindOptions | case insensitive search | named ranges | workbook scoped names | global named range | Cells.Find | search text in Excel | Aspose.Cells .NET
// Common Searches: search text in workbook scoped named ranges Aspose.Cells | case insensitive find in global named ranges C# | filter workbook scoped names Aspose.Cells | find all occurrences of a word in named ranges .NET | Aspose.Cells example search invoice
// Developer Intent: Find every cell containing the word “invoice” within all workbook‑scoped (global) named ranges, ignoring case.
// Use Cases: Verify that invoice identifiers are present in designated global ranges before report generation. | Create an audit log of cells that reference invoices across the workbook. | Collect cell addresses of invoice mentions for highlighting or downstream processing. | Automate validation of naming conventions for invoice data in Excel files.
// AI Prompts: How can I store the found cells in a List<Cell> instead of printing them? | Show me code to search for multiple keywords (e.g., "invoice", "receipt") across workbook‑scoped named ranges. | Explain how to stop after the first match per named range while keeping case‑insensitivity. | Provide a version that writes the results to a new worksheet.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, defines global named ranges, configures FindOptions for a case‑insensitive "contains" search, filters workbook‑scoped names, iterates each range, finds every "invoice" occurrence, writes results to the console, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Invoice #001");
            sheet.Cells["A2"].PutValue("No related text");
            sheet.Cells["B1"].PutValue("Some other data");
            sheet.Cells["C1"].PutValue("invoice details");
            sheet.Cells["D1"].PutValue("Random text");

            // Create workbook‑scoped (global) named ranges
            int idx1 = workbook.Worksheets.Names.Add("GlobalRange1");
            Name name1 = workbook.Worksheets.Names[idx1];
            name1.RefersTo = "=Sheet1!$A$1:$A$2";
            name1.SheetIndex = 0; // 0 = global scope

            int idx2 = workbook.Worksheets.Names.Add("GlobalRange2");
            Name name2 = workbook.Worksheets.Names[idx2];
            name2.RefersTo = "=Sheet1!$C$1:$C$1";
            name2.SheetIndex = 0; // global scope

            // Prepare case‑insensitive find options
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                CaseSensitive = false
            };

            // Retrieve all workbook‑scoped names
            Name[] globalNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);

            // Search each named range for the word "invoice"
            foreach (Name nm in globalNames)
            {
                // Get all ranges referred by this name
                AsposeRange[] ranges = nm.GetRanges();

                foreach (AsposeRange rng in ranges)
                {
                    // Define the search area corresponding to the current range
                    CellArea area = new CellArea
                    {
                        StartRow = rng.FirstRow,
                        StartColumn = rng.FirstColumn,
                        EndRow = rng.FirstRow + rng.RowCount - 1,
                        EndColumn = rng.FirstColumn + rng.ColumnCount - 1
                    };
                    findOptions.SetRange(area);

                    // Find first occurrence
                    Cell found = rng.Worksheet.Cells.Find("invoice", null, findOptions);

                    // Iterate through all occurrences within this range
                    while (found != null)
                    {
                        Console.WriteLine($"Found in named range '{nm.Text}' at cell {found.Name} (value: \"{found.StringValue}\")");
                        // Search for the next occurrence starting after the current cell
                        found = rng.Worksheet.Cells.Find("invoice", found, findOptions);
                    }
                }
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("InvoiceSearchResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
