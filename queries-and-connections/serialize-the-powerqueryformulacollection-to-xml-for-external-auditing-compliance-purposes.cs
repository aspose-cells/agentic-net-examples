using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

class PowerQueryFormulaXmlExporter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputXml = "PowerQueryFormulas.xml";

            // Verify input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input workbook '{inputPath}' was not found.");

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Safely obtain the Power Query formulas collection.
            PowerQueryFormulaCollection formulas = null;
            if (workbook.DataMashup != null)
                formulas = workbook.DataMashup.PowerQueryFormulas;

            // Create an XML writer to serialize the collection.
            using (XmlWriter writer = XmlWriter.Create(outputXml, new XmlWriterSettings { Indent = true }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("PowerQueryFormulas");

                if (formulas != null)
                {
                    foreach (PowerQueryFormula formula in formulas)
                    {
                        writer.WriteStartElement("PowerQueryFormula");

                        writer.WriteElementString("Name", formula.Name ?? string.Empty);
                        writer.WriteElementString("Description", formula.Description ?? string.Empty);
                        writer.WriteElementString("FormulaDefinition", formula.FormulaDefinition ?? string.Empty);
                        writer.WriteElementString("Type", formula.Type.ToString());

                        writer.WriteStartElement("Items");
                        foreach (PowerQueryFormulaItem item in formula.PowerQueryFormulaItems)
                        {
                            writer.WriteStartElement("Item");
                            writer.WriteElementString("Name", item.Name ?? string.Empty);
                            writer.WriteElementString("Value", item.Value ?? string.Empty);
                            writer.WriteEndElement(); // Item
                        }
                        writer.WriteEndElement(); // Items

                        writer.WriteEndElement(); // PowerQueryFormula
                    }
                }

                writer.WriteEndElement(); // PowerQueryFormulas
                writer.WriteEndDocument();
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}