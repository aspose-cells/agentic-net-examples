// Title: Generate an XML report of an Excel workbook's formula audit trail with timestamps using Aspose.Cells for .NET
// Common Searches: how to create a formula audit log with timestamps in Aspose.Cells C# | export Excel cell formulas and calculation results to XML using Aspose.Cells | Aspose.Cells C# generate XML report of workbook formulas after recalculation | C# retrieve formula, result and evaluation time from Excel with Aspose.Cells | save formula audit trail to XML file in .NET using Aspose.Cells
// Tags: Aspose.Cells export formula audit trail to XML | C# generate formula calculation log with timestamps | Excel workbook formula audit XML report | Aspose.Cells recalculate formulas and capture results | XML audit report for Excel formulas using .NET

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, forces a full formula recalculation, iterates through all worksheets and cells to collect each formula, its evaluated result, and a UTC timestamp, and writes these entries into an XML document (<FormulaAuditTrail>) that includes a generation timestamp. The XML file is saved as FormulaAuditTrail.xml, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Prepare the XML document for the audit trail
            XDocument xmlReport = new XDocument(
                new XElement("FormulaAuditTrail",
                    new XAttribute("GeneratedOn", DateTime.UtcNow.ToString("o")),
                    new XElement("Entries") // container for individual log entries
                )
            );

            XElement entriesElement = xmlReport.Root.Element("Entries");

            // Iterate through all worksheets and cells to collect formula information
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        string cellName = cell.Name;
                        string formula = cell.Formula;
                        string result = cell.Value?.ToString() ?? string.Empty;
                        string timestamp = DateTime.UtcNow.ToString("o");

                        entriesElement.Add(
                            new XElement("Entry",
                                new XElement("Worksheet", sheet.Name),
                                new XElement("Cell", cellName),
                                new XElement("Formula", formula),
                                new XElement("Result", result),
                                new XElement("Timestamp", timestamp)
                            )
                        );
                    }
                }
            }

            // Save the XML report to a file
            string xmlPath = "FormulaAuditTrail.xml";
            xmlReport.Save(xmlPath);
            Console.WriteLine($"Audit trail saved to '{xmlPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
