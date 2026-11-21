// Title: Export worksheet PowerQuery formulas to XML with Aspose.Cells in C#
// Common Searches: how can I extract Power Query formulas from an Excel file using Aspose.Cells C# | C# code to save PowerQueryFormulaCollection as an XML document for audit compliance | using reflection to access PowerQueryFormulaCollection in Aspose.Cells | serialize Power Query formula names and expressions to XML with Aspose.Cells .NET | export Power Query formulas to XML for compliance reporting in C#
// Tags: Aspose.Cells PowerQueryFormulaCollection XML export | C# reflection Power Query formula extraction | Excel audit compliance Power Query XML generation | serialize worksheet Power Query formulas

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, uses reflection to obtain the worksheet's PowerQueryFormulaCollection, iterates through each formula to read its Name and Formula properties, builds an XML document with <Formula> elements containing these attributes, creates the output folder if needed, and writes the XML file for external auditing or compliance purposes.
class PowerQueryFormulaSerializer
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "PowerQueryFormulas.xml";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare XML document
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement root = xmlDoc.CreateElement("PowerQueryFormulas");
            xmlDoc.AppendChild(root);

            // Use reflection to access PowerQueryFormulaCollection (avoids compile‑time dependency)
            var pqCollectionProp = worksheet.GetType().GetProperty("PowerQueryFormulaCollection");
            if (pqCollectionProp != null)
            {
                var pqCollection = pqCollectionProp.GetValue(worksheet, null);
                if (pqCollection != null)
                {
                    var enumerable = pqCollection as System.Collections.IEnumerable;
                    if (enumerable != null)
                    {
                        foreach (var pqFormula in enumerable)
                        {
                            if (pqFormula == null) continue;

                            var nameProp = pqFormula.GetType().GetProperty("Name");
                            var formulaProp = pqFormula.GetType().GetProperty("Formula");

                            string name = nameProp?.GetValue(pqFormula)?.ToString() ?? string.Empty;
                            string formula = formulaProp?.GetValue(pqFormula)?.ToString() ?? string.Empty;

                            XmlElement formulaElem = xmlDoc.CreateElement("Formula");
                            formulaElem.SetAttribute("Name", name);
                            formulaElem.SetAttribute("Formula", formula);
                            root.AppendChild(formulaElem);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Power Query formulas are not supported in this version of Aspose.Cells.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the XML document
            xmlDoc.Save(outputPath);
            Console.WriteLine($"PowerQuery formulas have been serialized to \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
