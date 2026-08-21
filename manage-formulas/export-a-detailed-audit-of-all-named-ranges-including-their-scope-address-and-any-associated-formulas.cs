// Title: Audit all named ranges in an Excel workbook – scope, address, and RefersTo using Aspose.Cells for .NET
// Description: Loads a workbook, creates a "NamedRangeAudit" sheet, lists every defined name with its scope (workbook or worksheet), resolved address, and original RefersTo formula, auto‑fits columns, and saves the updated file.
// Keywords: Aspose.Cells | C# | .NET | named range audit | list defined names | named range scope | range address | RefersTo formula | Excel workbook report
// Common Searches: list all named ranges Aspose.Cells C# | export named range details to new worksheet | get named range address Aspose.Cells .NET | audit named ranges in Excel using Aspose | retrieve named range scope programmatically
// Developer Intent: Create a worksheet that enumerates every named range in a workbook together with its scope, resolved address, and reference formula.
// Use Cases: Produce a compliance report of all named ranges before distributing a workbook. | Locate and troubleshoot named ranges that point to incorrect or external cells. | Provide end‑users with a summary sheet that explains each named range’s purpose and location.
// AI Prompts: Generate C# code with Aspose.Cells that writes a report of all defined names, including scope, address, and RefersTo, to a new worksheet. | Show how to safely call Name.GetRange() and fall back to the raw RefersTo string when the range cannot be resolved. | Explain how to determine whether a named range is workbook‑level or worksheet‑level using the Name.SheetIndex property.

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, creates a "NamedRangeAudit" sheet, lists every defined name with its scope (workbook or worksheet), resolved address, and original RefersTo formula, auto‑fits columns, and saves the updated file.
class NamedRangeAudit
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output_with_audit.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the audit report
            int auditSheetIndex = workbook.Worksheets.Add();
            Worksheet auditSheet = workbook.Worksheets[auditSheetIndex];
            auditSheet.Name = "NamedRangeAudit";

            // Write header row
            Cells auditCells = auditSheet.Cells;
            auditCells["A1"].PutValue("Name");
            auditCells["B1"].PutValue("Scope");
            auditCells["C1"].PutValue("Address");
            auditCells["D1"].PutValue("RefersTo");

            int currentRow = 1; // Zero‑based index; row 1 is the second row

            // Iterate through all defined names in the workbook
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // Determine the scope of the name
                string scope = definedName.SheetIndex == 0
                    ? "Workbook"
                    : $"Worksheet: {workbook.Worksheets[definedName.SheetIndex - 1].Name}";

                // Try to obtain the actual range address
                string address = string.Empty;
                try
                {
                    Aspose.Cells.Range range = definedName.GetRange();
                    if (range != null)
                    {
                        address = range.Address;
                    }
                }
                catch
                {
                    // GetRange may throw if the name does not refer to a range; ignore
                }

                // Fallback to the raw RefersTo string if address could not be resolved
                if (string.IsNullOrEmpty(address))
                {
                    address = definedName.RefersTo?.TrimStart('=');
                }

                // Populate the audit row
                auditCells[currentRow, 0].PutValue(definedName.Text);          // Name
                auditCells[currentRow, 1].PutValue(scope);                   // Scope
                auditCells[currentRow, 2].PutValue(address);                // Address
                auditCells[currentRow, 3].PutValue(definedName.RefersTo);    // Formula/reference

                currentRow++;
            }

            // Adjust column widths for readability
            auditSheet.AutoFitColumns();

            // Save the workbook with the audit sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Audit completed. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
