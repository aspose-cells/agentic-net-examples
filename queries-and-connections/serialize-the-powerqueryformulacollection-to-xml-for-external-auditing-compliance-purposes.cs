// Title: How to export Power Query formulas from an Excel workbook to an XML file with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, finds every cell whose formula contains the keyword POWERQUERY, and writes each formula into an XML document. | Update the XML generation so that each <Formula> element also includes the originating cell address (e.g., A1) as an attribute. | Add robust error handling and console logging that reports the number of Power Query formulas found in each worksheet. | Create a reusable method that accepts a workbook path and returns an XDocument containing all Power Query formulas ready for saving.
// Common Searches: c# Aspose.Cells extract Power Query formulas from Excel workbook | save Power Query M code to XML using Aspose.Cells .NET | list cells containing POWERQUERY formula in .xlsx file | generate XML audit report of Power Query formulas in Excel with C# | how to count Power Query formulas per worksheet using Aspose.Cells
// Tags: Aspose.Cells export Power Query formulas to XML | C# serialize Power Query M code | Excel workbook Power Query formula extraction | XML audit report for Power Query formulas | cell address inclusion in Power Query XML output | error handling for Power Query extraction Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Xml.Linq;

// The program loads an Excel workbook with Aspose.Cells, scans all cells for formulas containing the keyword POWERQUERY, aggregates those formulas (optionally with their cell addresses) into an XDocument, and saves the result as an XML file for auditing purposes.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputXml = "PowerQueryFormulas.xml";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains Power Query formulas.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust the index if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Collect all Power Query formulas from the worksheet.
            var pqFormulas = new System.Collections.Generic.List<string>();
            Cells cells = worksheet.Cells;
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) &&
                    cell.Formula.IndexOf("POWERQUERY", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    pqFormulas.Add(cell.Formula);
                }
            }

            // Create an XML document to store the collected formulas.
            XDocument doc = new XDocument(
                new XElement("PowerQueryFormulas",
                    new XElement("Count", pqFormulas.Count),
                    new XElement("Formulas",
                        new System.Collections.Generic.List<XElement>(
                            pqFormulas.ConvertAll(f => new XElement("Formula", f))
                        )
                    )
                )
            );

            // Save the XML document.
            doc.Save(outputXml);
            Console.WriteLine($"Power Query formulas have been written to \"{outputXml}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
