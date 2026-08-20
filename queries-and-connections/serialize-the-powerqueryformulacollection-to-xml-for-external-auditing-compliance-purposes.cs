// Title: C# – Export PowerQueryFormulaCollection to XML with Aspose.Cells for Audit Compliance
// Description: Loads an Excel workbook, accesses its DataMashup, iterates the PowerQueryFormulaCollection, builds a structured XML document with formula attributes and items, and saves the file for external audit and governance purposes.
// Keywords: Aspose.Cells | PowerQueryFormulaCollection | C# XML export | DataMashup serialization | audit compliance | Excel Power Query export | XML audit file | Power Query governance | Aspose.Cells .NET example
// Common Searches: export PowerQuery formulas to XML Aspose.Cells | serialize PowerQueryFormulaCollection C# | create audit XML from Excel Power Query | Aspose.Cells DataMashup XML output | C# code to save Power Query definitions as XML
// Developer Intent: Generate an XML file that captures every Power Query formula and its items from an Excel workbook for compliance auditing.
// Use Cases: Produce a regulatory audit report of Power Query transformations. | Create version‑controlled XML snapshots to compare workbook changes. | Feed exported XML into governance tools for automated policy validation.
// AI Prompts: Write C# code using Aspose.Cells to export a PowerQueryFormulaCollection to XML with attributes Name, Description, Type, FormulaDefinition and child Item elements. | Add comprehensive error handling and logging for missing DataMashup, file‑system errors, and unexpected null values in the Power Query export sample. | Show how to deserialize the generated PowerQueryFormulasAudit.xml back into a PowerQueryFormulaCollection using Aspose.Cells.

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace PowerQueryFormulaXmlExport
{
    // Loads an Excel workbook, accesses its DataMashup, iterates the PowerQueryFormulaCollection, builds a structured XML document with formula attributes and items, and saves the file for external audit and governance purposes.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook that contains Power Query formulas
            string sourcePath = "SourceWithPowerQuery.xlsx";

            // Verify that the source file exists before attempting to load it
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Access the DataMashup object which holds the PowerQueryFormulaCollection
                DataMashup mashup = workbook.DataMashup;

                // Guard against null DataMashup (should not happen for a valid workbook)
                if (mashup == null)
                {
                    Console.WriteLine("The workbook does not contain mashup data.");
                    return;
                }

                // Retrieve the collection of Power Query formulas
                PowerQueryFormulaCollection formulas = mashup.PowerQueryFormulas;

                // Create the root XML element
                XElement root = new XElement("PowerQueryFormulas");

                // Iterate through each formula and build XML representation
                foreach (PowerQueryFormula formula in formulas)
                {
                    // Create an element for the formula with its main properties as attributes
                    XElement formulaElement = new XElement("PowerQueryFormula",
                        new XAttribute("Name", formula.Name ?? string.Empty),
                        new XAttribute("Description", formula.Description ?? string.Empty),
                        new XAttribute("Type", formula.Type.ToString()),
                        new XAttribute("FormulaDefinition", formula.FormulaDefinition ?? string.Empty)
                    );

                    // Add child elements for each formula item (if any)
                    PowerQueryFormulaItemCollection items = formula.PowerQueryFormulaItems;
                    foreach (PowerQueryFormulaItem item in items)
                    {
                        XElement itemElement = new XElement("Item",
                            new XAttribute("Name", item.Name ?? string.Empty),
                            new XAttribute("Value", item.Value ?? string.Empty)
                        );
                        formulaElement.Add(itemElement);
                    }

                    // Append the formula element to the root
                    root.Add(formulaElement);
                }

                // Build the final XML document
                XDocument doc = new XDocument(
                    new XDeclaration("1.0", "utf-8", "yes"),
                    root
                );

                // Define the output XML file path
                string xmlOutputPath = "PowerQueryFormulasAudit.xml";

                // Save the XML document to file
                doc.Save(xmlOutputPath);

                Console.WriteLine($"Power Query formulas have been serialized to XML at: {Path.GetFullPath(xmlOutputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
