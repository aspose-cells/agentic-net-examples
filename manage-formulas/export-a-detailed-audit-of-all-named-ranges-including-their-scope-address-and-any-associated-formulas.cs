// Title: Audit All Named Ranges (Scope, Address, Formula) with Aspose.Cells for .NET
// Description: C# sample that enumerates every defined name in a workbook, identifies its scope (global or worksheet), shows the RefersTo expression, resolves the range address when possible, and writes the details to a text file.
// Keywords: Aspose.Cells | C# named range audit | list named ranges .NET | global vs local named range | GetRange Aspose.Cells | export named ranges to file | Excel named range scope
// Common Searches: Aspose.Cells list all named ranges | How to get named range address in C# | Export named range definitions to text file | Determine scope of named ranges using Aspose.Cells | Retrieve formula of a named range .NET
// Developer Intent: Produce a text report that lists each defined name, its scope, RefersTo string, and resolved range address.
// Use Cases: Generate compliance documentation of named ranges in automatically created workbooks | Validate that worksheet‑level names reference the intended cells before publishing | Debug complex spreadsheets by logging names, formulas, and addresses | Create a reference file for auditors describing workbook naming conventions
// AI Prompts: Write C# code with Aspose.Cells that iterates over Workbook.Worksheets.Names and outputs Name, Scope, RefersTo, and Address to a CSV file. | Provide a method returning a collection of objects (Name, Scope, RefersTo, Address) for all defined names, handling non‑range names without throwing. | Explain how Aspose.Cells distinguishes global and local named ranges and how to retrieve the worksheet index for a local name.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// C# sample that enumerates every defined name in a workbook, identifies its scope (global or worksheet), shows the RefersTo expression, resolves the range address when possible, and writes the details to a text file.
class NamedRangeAudit
{
    static void Main()
    {
        try
        {
            // Create a new workbook (replace with Workbook wb = new Workbook("input.xlsx") for an existing file)
            Workbook wb = new Workbook();

            // -------------------------------------------------
            // Sample data – this section can be removed when using a real workbook
            // -------------------------------------------------
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Sheet1";
            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].PutValue(30);
            ws.Cells["B1"].Formula = "=SUM(A1:A3)";

            // Global named range
            int globalIdx = wb.Worksheets.Names.Add("GlobalRange");
            Name globalName = wb.Worksheets.Names[globalIdx];
            globalName.RefersTo = "=Sheet1!$A$1:$A$3";

            // Worksheet‑level named range (belongs to the first sheet, index is 1‑based)
            int sheetIdx = wb.Worksheets.Names.Add("SheetRange");
            Name sheetName = wb.Worksheets.Names[sheetIdx];
            sheetName.RefersTo = "=Sheet1!$B$1";
            sheetName.SheetIndex = 1; // makes it local to the first worksheet
            // -------------------------------------------------
            // End of sample data
            // -------------------------------------------------

            // Prepare a text file to store the audit
            using (StreamWriter writer = new StreamWriter("NamedRangeAudit.txt"))
            {
                // Iterate through every defined name in the workbook
                foreach (Name name in wb.Worksheets.Names)
                {
                    // Determine the scope of the name
                    string scope = name.SheetIndex == 0
                        ? "Workbook (Global)"
                        : $"Worksheet (Index {name.SheetIndex - 1})";

                    // The raw RefersTo string (e.g., "=Sheet1!$A$1:$A$3" or a formula)
                    string refersTo = name.RefersTo;

                    // Try to obtain a concrete range address if the name actually points to a range
                    string rangeAddress = string.Empty;
                    try
                    {
                        AsposeRange rng = name.GetRange();
                        if (rng != null)
                        {
                            rangeAddress = rng.Address;
                        }
                    }
                    catch
                    {
                        // GetRange throws if the name does not refer to a range – ignore
                    }

                    // Build the audit line
                    string line = $"Name: {name.Text}, Scope: {scope}, RefersTo: {refersTo}";
                    if (!string.IsNullOrEmpty(rangeAddress))
                    {
                        line += $", Range Address: {rangeAddress}";
                    }

                    // Output to console and to the file
                    Console.WriteLine(line);
                    writer.WriteLine(line);
                }
            }

            // Save the workbook (optional – demonstrates the required save lifecycle)
            wb.Save("AuditWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
