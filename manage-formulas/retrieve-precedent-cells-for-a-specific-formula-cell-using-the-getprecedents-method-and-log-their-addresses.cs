// Title: Aspose.Cells C# GetPrecedents – List and Log Formula Cell Dependencies (including external links)
// Description: Creates a workbook, assigns a formula to A1 that references internal cells, a range, and an external workbook, then uses GetPrecedents to retrieve a ReferredAreaCollection. Each precedent is formatted with optional external file name, sheet name, and start/end addresses, printed to the console, and the workbook is saved.
// Keywords: Aspose.Cells | GetPrecedents | C# | .NET | formula precedents | cell dependency | external workbook reference | ReferredAreaCollection | cell address extraction | Excel formula audit
// Common Searches: Aspose.Cells GetPrecedents C# example | how to list precedent cells of a formula using Aspose.Cells | retrieve external links from Excel formula with Aspose.Cells | C# code to get cell dependencies in Aspose.Cells | enumerate formula precedents in .NET
// Developer Intent: Extract every cell or range that a formula depends on and output their full addresses.
// Use Cases: Audit a worksheet by enumerating all cells and ranges referenced by a specific formula, including links to other workbooks. | Generate a dependency report that shows which sheets and external files are used by a formula for impact analysis. | Validate complex formulas programmatically by confirming that all referenced cells and ranges are correct.
// AI Prompts: Write C# code with Aspose.Cells that calls GetPrecedents on a formula cell, formats each precedent with sheet name and external file name when present, and prints the addresses. | Show how to handle both single‑cell and range precedents returned by GetPrecedents, including external workbook references, in a .NET example. | Provide a snippet that extracts precedent information, logs it to the console, and then saves the workbook, indicating the output file path.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, assigns a formula to A1 that references internal cells, a range, and an external workbook, then uses GetPrecedents to retrieve a ReferredAreaCollection. Each precedent is formatted with optional external file name, sheet name, and start/end addresses, printed to the console, and the workbook is saved.
    public class GetPrecedentsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define a formula that references several cells/ranges, including an external link
                cells["A1"].Formula = "=B1+SUM(C1:D2)+[External.xlsx]Sheet1!E5";

                // Retrieve all precedent references for the formula cell
                ReferredAreaCollection precedents = cells["A1"].GetPrecedents();

                if (precedents != null)
                {
                    foreach (ReferredArea area in precedents)
                    {
                        StringBuilder sb = new StringBuilder();

                        // Include external file name if the reference is external
                        if (area.IsExternalLink)
                        {
                            sb.Append('[').Append(area.ExternalFileName).Append(']');
                        }

                        // Append sheet name and start cell address
                        sb.Append(area.SheetName).Append('!');
                        sb.Append(CellsHelper.CellIndexToName(area.StartRow, area.StartColumn));

                        // If the reference is a range, append the end cell address
                        if (area.IsArea)
                        {
                            sb.Append(':').Append(CellsHelper.CellIndexToName(area.EndRow, area.EndColumn));
                        }

                        // Log the full address of the precedent
                        Console.WriteLine(sb.ToString());
                    }
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "GetPrecedentsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            GetPrecedentsDemo.Run();
        }
    }
}
