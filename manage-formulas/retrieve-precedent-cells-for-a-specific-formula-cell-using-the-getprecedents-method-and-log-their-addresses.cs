// Title: Get Precedent Cell Addresses for a Formula with Aspose.Cells (C#)
// Description: Creates a workbook, sets a formula in A1 that references internal cells, a range, and an external workbook, then uses the GetPrecedents method to enumerate each ReferredArea, builds a full address (including external file name, sheet name, and range limits), writes the addresses to the console, and saves the file.
// Keywords: Aspose.Cells | GetPrecedents | C# | .NET | precedent cells | formula dependencies | external link reference | ReferredArea | cell address range | enumerate precedents
// Common Searches: Aspose.Cells GetPrecedents C# example | how to list precedent cells of a formula in .NET | retrieve external references from a formula using Aspose.Cells | enumerate precedent ranges in Excel with Aspose.Cells | C# code to get dependent cells of a formula
// Developer Intent: Obtain every cell or range that a formula depends on, format each reference with sheet and external file information, and output the list programmatically.
// Use Cases: Audit formula dependencies before restructuring a workbook to avoid breaking calculations. | Generate a dependency report that lists all internal and external cells influencing a key metric. | Validate external workbook links for data‑integrity checks in automated spreadsheet processing.
// AI Prompts: Write C# code using Aspose.Cells to retrieve all precedent cells for a given formula and print each address with sheet and external file names. | Create a helper method that converts ReferredArea objects from GetPrecedents into readable strings handling single cells, ranges, and external links. | Explain how to iterate over the ReferredAreaCollection returned by GetPrecedents and export the addresses to a CSV file.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsPrecedentsDemo
{
    // Creates a workbook, sets a formula in A1 that references internal cells, a range, and an external workbook, then uses the GetPrecedents method to enumerate each ReferredArea, builds a full address (including external file name, sheet name, and range limits), writes the addresses to the console, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Define a formula that references several cells, a range, and an external link
            cells["A1"].Formula = "=B1+SUM(B2:B5)+[Book1.xls]Sheet1!C3";

            // Retrieve all precedent references of the formula cell A1
            ReferredAreaCollection precedents = cells["A1"].GetPrecedents();

            // If there are precedents, iterate and log their addresses
            if (precedents != null && precedents.Count > 0)
            {
                Console.WriteLine("Precedent references for cell A1:");
                foreach (ReferredArea area in precedents)
                {
                    StringBuilder sb = new StringBuilder();

                    // Include external file name if the reference is an external link
                    if (area.IsExternalLink)
                    {
                        sb.Append($"[{area.ExternalFileName}]");
                    }

                    // Append sheet name
                    sb.Append($"{area.SheetName}!");

                    // Append start cell address
                    sb.Append(CellsHelper.CellIndexToName(area.StartRow, area.StartColumn));

                    // If the reference is an area (range), append the end cell address
                    if (area.IsArea)
                    {
                        sb.Append($":{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}");
                    }

                    // Output the constructed address
                    Console.WriteLine(sb.ToString());
                }
            }
            else
            {
                Console.WriteLine("No precedents found for cell A1.");
            }

            // Optionally save the workbook (demonstrates lifecycle rule usage)
            workbook.Save("PrecedentsDemo.xlsx");
        }
    }
}
