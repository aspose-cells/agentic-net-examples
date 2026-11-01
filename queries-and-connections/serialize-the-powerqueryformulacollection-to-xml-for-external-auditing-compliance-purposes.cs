// Title: Export Power Query formulas from an Excel worksheet to XML using Aspose.Cells for .NET compliance auditing
// AI Prompts: Generate C# code that loads a workbook with Aspose.Cells, accesses the worksheet's PowerQueryFormulaCollection via reflection, and writes each formula's Name, Formula, and IsEnabled attributes into an XML document. | Create a version‑agnostic .NET routine that extracts all Power Query definitions from a worksheet and saves them as <PowerQueryFormula> elements in an XML file for external audit purposes.
// Common Searches: how to extract Power Query formulas from an Excel file using Aspose.Cells C# | serialize PowerQueryFormulaCollection to XML for compliance reporting | C# reflection example to read Power Query definitions in Aspose.Cells | export Power Query formula name and expression to XML with Aspose.Cells | audit Excel Power Query queries programmatically .NET
// Tags: Aspose.Cells Power Query XML export | C# PowerQueryFormulaCollection serialization | use reflection to get Power Query collection Aspose.Cells | audit Excel Power Query definitions .NET | XML compliance output for Power Query formulas

using Aspose.Cells;
using System;
using System.IO;
using System.Xml.Linq;
using System.Collections;

// The sample loads an Excel workbook with Aspose.Cells, uses reflection to obtain the worksheet's PowerQueryFormulaCollection, iterates through each PowerQueryFormula extracting Name, Formula, and IsEnabled, builds an XDocument containing <PowerQueryFormula> elements with these attributes, and saves the XML file for external compliance auditing while handling missing APIs and runtime errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputXmlPath = "PowerQueryFormulas.xml";

            // Ensure the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that may contain Power Query formulas.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Attempt to retrieve the PowerQueryFormulaCollection via reflection
            // (the API may not be available in older Aspose.Cells versions).
            var pqCollectionObj = worksheet.GetType()
                                           .GetProperty("PowerQueryFormulaCollection")
                                           ?.GetValue(worksheet);

            // Prepare the XML document that will hold the serialized data.
            XDocument xmlDoc = new XDocument(new XElement("PowerQueryFormulas"));

            if (pqCollectionObj is IEnumerable pqCollection)
            {
                // Iterate through each PowerQueryFormula (using reflection to stay version‑agnostic).
                foreach (var pqFormula in pqCollection)
                {
                    var formulaType = pqFormula.GetType();

                    string name = formulaType.GetProperty("Name")?.GetValue(pqFormula) as string ?? string.Empty;
                    string formula = formulaType.GetProperty("Formula")?.GetValue(pqFormula) as string ?? string.Empty;
                    bool isEnabled = false;
                    var isEnabledProp = formulaType.GetProperty("IsEnabled");
                    if (isEnabledProp != null && isEnabledProp.GetValue(pqFormula) is bool enabled)
                    {
                        isEnabled = enabled;
                    }

                    // Add the formula details to the XML.
                    XElement formulaElement = new XElement("PowerQueryFormula",
                        new XAttribute("Name", name),
                        new XAttribute("Formula", formula),
                        new XAttribute("IsEnabled", isEnabled));

                    xmlDoc.Root.Add(formulaElement);
                }
            }
            else
            {
                Console.WriteLine("No Power Query formulas found or the API is unavailable in this Aspose.Cells version.");
            }

            // Save the XML document to a file for external auditing.
            xmlDoc.Save(outputXmlPath);
            Console.WriteLine($"Power Query formulas exported to '{outputXmlPath}'.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors without crashing the application.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
