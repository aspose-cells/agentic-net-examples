using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace PowerQueryFormulaXmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source workbook that contains Power Query formulas
                string workbookPath = "source_with_powerquery.xlsx";

                // Verify that the workbook file exists before attempting to load it
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                    return;
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(workbookPath);

                // Access the DataMashup object which holds the Power Query formulas collection
                var mashup = workbook.DataMashup;
                var formulas = mashup?.PowerQueryFormulas; // use var/dynamic to avoid compile‑time type dependency

                // Prepare an XML document to hold the serialized data
                XmlDocument xmlDoc = new XmlDocument();

                // Create XML declaration
                XmlDeclaration xmlDecl = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
                xmlDoc.AppendChild(xmlDecl);

                // Root element
                XmlElement root = xmlDoc.CreateElement("PowerQueryFormulas");
                xmlDoc.AppendChild(root);

                if (formulas != null)
                {
                    // Iterate through each PowerQueryFormula using dynamic typing
                    foreach (dynamic formula in formulas)
                    {
                        XmlElement formulaElem = xmlDoc.CreateElement("PowerQueryFormula");
                        root.AppendChild(formulaElem);

                        AppendTextElement(xmlDoc, formulaElem, "Name", formula.Name);
                        AppendTextElement(xmlDoc, formulaElem, "Description", formula.Description);
                        AppendTextElement(xmlDoc, formulaElem, "Type", formula.Type?.ToString());
                        AppendTextElement(xmlDoc, formulaElem, "FormulaDefinition", formula.FormulaDefinition);

                        // Serialize the collection of items belonging to the formula
                        var items = formula.PowerQueryFormulaItems;
                        XmlElement itemsElem = xmlDoc.CreateElement("Items");
                        formulaElem.AppendChild(itemsElem);

                        if (items != null)
                        {
                            foreach (dynamic item in items)
                            {
                                XmlElement itemElem = xmlDoc.CreateElement("Item");
                                itemsElem.AppendChild(itemElem);

                                AppendTextElement(xmlDoc, itemElem, "Name", item.Name);
                                AppendTextElement(xmlDoc, itemElem, "Value", item.Value);
                            }
                        }
                    }
                }
                else
                {
                    // No formulas found – add a placeholder comment node for clarity
                    XmlComment comment = xmlDoc.CreateComment("No Power Query formulas found in the workbook.");
                    root.AppendChild(comment);
                }

                // Save the XML document to a file
                string xmlOutputPath = "PowerQueryFormulasAudit.xml";
                xmlDoc.Save(xmlOutputPath);

                Console.WriteLine($"Power Query formulas have been serialized to '{xmlOutputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <summary>
        /// Helper method to create an element with inner text and append it to a parent node.
        /// </summary>
        private static void AppendTextElement(XmlDocument doc, XmlElement parent, string elementName, string innerText)
        {
            XmlElement elem = doc.CreateElement(elementName);
            elem.InnerText = innerText ?? string.Empty;
            parent.AppendChild(elem);
        }
    }
}