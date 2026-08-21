// Title: Aspose.Cells .NET – Read IsAutomaticPaperSize for All Worksheets into a Dictionary
// Description: Load a workbook, loop through every worksheet, retrieve the PageSetup.IsAutomaticPaperSize flag, store each boolean in a Dictionary keyed by sheet name, and optionally save the file.
// Keywords: Aspose.Cells | C# | .NET | IsAutomaticPaperSize | PageSetup | worksheet iteration | dictionary storage | workbook printing settings | Excel automation
// Common Searches: Aspose.Cells get IsAutomaticPaperSize for each sheet | C# read worksheet page setup flag | store worksheet properties in a dictionary Aspose | iterate worksheets and access PageSetup in .NET | how to check automatic paper size in Excel using Aspose
// Developer Intent: Extract the IsAutomaticPaperSize flag from every worksheet and keep the results in a name‑based dictionary.
// Use Cases: Audit a workbook to ensure all sheets are set to automatic paper size before batch printing. | Create a configuration report that lists page‑setup flags for documentation or compliance. | Programmatically adjust page settings only on sheets where automatic sizing is disabled.
// AI Prompts: Generate C# code with Aspose.Cells that returns a Dictionary<string, bool> containing each worksheet’s IsAutomaticPaperSize value. | Show how to log the IsAutomaticPaperSize flag for all sheets and save the workbook only when at least one sheet has the flag set to false. | Explain how to toggle IsAutomaticPaperSize for selected worksheets while preserving other PageSetup properties.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Load a workbook, loop through every worksheet, retrieve the PageSetup.IsAutomaticPaperSize flag, store each boolean in a Dictionary keyed by sheet name, and optionally save the file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (provide the correct file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to store the IsAutomaticPaperSize value for each worksheet
        Dictionary<string, bool> automaticPaperSizeBySheet = new Dictionary<string, bool>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Read the IsAutomaticPaperSize property from the worksheet's PageSetup
            bool isAutomatic = sheet.PageSetup.IsAutomaticPaperSize;

            // Store the result using the worksheet name as the key
            automaticPaperSizeBySheet[sheet.Name] = isAutomatic;

            // Output the value for verification
            Console.WriteLine($"Worksheet '{sheet.Name}': IsAutomaticPaperSize = {isAutomatic}");
        }

        // Save the workbook if any changes were made (optional)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
