using System;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace PowerQueryFormulaXmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook that contains Power Query formulas
            string sourcePath = "SourceWithPowerQuery.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the Power Query formulas collection via DataMashup
            var formulas = workbook.DataMashup?.PowerQueryFormulas;

            if (formulas == null || formulas.Count == 0)
            {
                Console.WriteLine("No Power Query formulas found in the workbook.");
                return;
            }

            // Create the root element for the XML document
            XElement root = new XElement("PowerQueryFormulas");

            // Iterate through each PowerQueryFormula and serialize its properties
            foreach (PowerQueryFormula formula in formulas)
            {
                XElement formulaElement = new XElement("PowerQueryFormula",
                    new XAttribute("Name", formula.Name ?? string.Empty),
                    new XAttribute("Type", formula.Type.ToString()),
                    new XAttribute("Description", formula.Description ?? string.Empty),
                    new XElement("FormulaDefinition", formula.FormulaDefinition ?? string.Empty)
                );

                // Serialize the collection of items belonging to the formula
                PowerQueryFormulaItemCollection items = formula.PowerQueryFormulaItems;
                if (items != null && items.Count > 0)
                {
                    XElement itemsElement = new XElement("Items");
                    foreach (PowerQueryFormulaItem item in items)
                    {
                        XElement itemElement = new XElement("Item",
                            new XAttribute("Name", item.Name ?? string.Empty),
                            new XAttribute("Value", item.Value ?? string.Empty)
                        );
                        itemsElement.Add(itemElement);
                    }
                    formulaElement.Add(itemsElement);
                }

                root.Add(formulaElement);
            }

            // Build the XDocument and save it to an XML file
            XDocument doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
            string outputXmlPath = "PowerQueryFormulasAudit.xml";
            doc.Save(outputXmlPath);

            Console.WriteLine($"Power Query formulas have been serialized to '{outputXmlPath}'.");
        }
    }
}