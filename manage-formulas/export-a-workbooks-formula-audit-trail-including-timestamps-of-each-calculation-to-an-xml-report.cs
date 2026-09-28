// Title: Export a workbook’s formula audit trail with calculation timestamps to an XML file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, forces a full formula recalculation, and creates an XML document listing each formula’s sheet name, cell address, expression, evaluated value, and the timestamp of calculation. | Develop a .NET method that scans all worksheets in a workbook, captures every formula cell’s details together with the current UTC time, and saves the collection as a structured XML report via Aspose.Cells.
// Common Searches: Aspose.Cells C# export all formulas with timestamps to XML | How to generate a formula audit XML report from an Excel file using Aspose.Cells | C# retrieve formula expression, value and calculation date from workbook | Save Excel formula audit trail as XML using Aspose.Cells for .NET | Export workbook formula audit log with timestamps in C#
// Tags: export formula audit trail to XML with Aspose.Cells | extract cell formulas and timestamps C# | calculate workbook formulas before XML export | Aspose.Cells generate formula audit XML report | C# workbook formula audit logging

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// // Loads an existing Excel workbook, forces a full formula recalculation, iterates through every worksheet and cell to collect each formula's sheet name, address, expression, evaluated value, and the current timestamp, then writes these details into an XML document named FormulaAuditReport.xml.
class FormulaAuditExport
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            const string inputPath = "input.xlsx";
            const string outputPath = "FormulaAuditReport.xml";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found – {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure all formulas are calculated
            workbook.CalculateFormula();

            // Root element for the XML report
            XElement root = new XElement("FormulaAuditTrail");

            // Iterate through all worksheets and cells to collect formula information
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        // Build an XML element for each formula cell
                        XElement formulaElement = new XElement("Formula",
                            new XAttribute("Sheet", sheet.Name),
                            new XAttribute("Address", cell.Name),
                            new XAttribute("Expression", cell.Formula),
                            new XAttribute("Value", cell.Value?.ToString() ?? string.Empty),
                            // Use the current time as the calculation timestamp
                            new XAttribute("CalculatedOn", DateTime.Now.ToString("o"))
                        );

                        root.Add(formulaElement);
                    }
                }
            }

            // Create the final XML document
            XDocument auditReport = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                root
            );

            // Save the XML report
            auditReport.Save(outputPath);
            Console.WriteLine($"Formula audit report saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
