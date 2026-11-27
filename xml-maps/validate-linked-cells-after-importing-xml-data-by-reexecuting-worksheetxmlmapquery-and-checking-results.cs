// Title: Validate linked cells after XML import with Worksheet.GetAreasOfXmlMapQuery in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# verify cells linked to an XML map after ImportXml | GetAreasOfXmlMapQuery example to check imported XML values in a worksheet | C# code to compare expected XML items with worksheet cell contents using Aspose.Cells | Validate XML map query results programmatically in Aspose.Cells .NET | How to ensure linked cells contain correct data after importing XML with Aspose.Cells
// Tags: Worksheet.GetAreasOfXmlMapQuery C# usage | Aspose.Cells XML map cell validation | ImportXml linked cells extraction | expected XML items verification in worksheet | XML map query result validation .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample creates a temporary XML file, imports it into a new workbook via Workbook.ImportXml, retrieves the generated XmlMap, uses Worksheet.GetAreasOfXmlMapQuery to obtain the cell areas linked to the '/Root/Data/Item' path, extracts the cell values, compares them against the expected list ['Alpha','Beta','Gamma'], reports whether validation succeeded, and optionally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // --------------------------- 1. Prepare sample XML ---------------------------
            // The XML schema will be inferred automatically by ImportXml.
            // This XML contains a root element with a collection of <Item> elements.
            string xmlContent = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Data>
        <Item>Alpha</Item>
        <Item>Beta</Item>
        <Item>Gamma</Item>
    </Data>
</Root>";

            // --------------------------- 2. Write XML to a temporary file ---------------------------
            string tempXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xml");
            try
            {
                File.WriteAllText(tempXmlPath, xmlContent);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to write temporary XML file: " + ex.Message);
                return;
            }

            // --------------------------- 3. Create workbook and import XML ---------------------------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Sheet1";

            try
            {
                if (!File.Exists(tempXmlPath))
                    throw new FileNotFoundException("Temporary XML file not found.", tempXmlPath);

                // ImportXml creates an XmlMap and populates the worksheet.
                workbook.ImportXml(tempXmlPath, "Sheet1", 0, 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to import XML: " + ex.Message);
                return;
            }
            finally
            {
                // Clean up the temporary file.
                try { if (File.Exists(tempXmlPath)) File.Delete(tempXmlPath); } catch { }
            }

            // --------------------------- 4. Retrieve the XmlMap ---------------------------
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XmlMap was created after ImportXml.");
                return;
            }

            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
            Console.WriteLine($"XmlMap created. Name: {xmlMap.Name}");

            // --------------------------- 5. Query linked cells ---------------------------
            // Path to the <Item> elements.
            string queryPath = "/Root/Data/Item";

            CellArea[] linkedAreas;
            try
            {
                linkedAreas = worksheet.GetAreasOfXmlMapQuery(queryPath, xmlMap);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while executing GetAreasOfXmlMapQuery: " + ex.Message);
                return;
            }

            if (linkedAreas == null || linkedAreas.Length == 0)
            {
                Console.WriteLine($"No cells are linked to the path '{queryPath}'.");
                return;
            }

            Console.WriteLine($"Found {linkedAreas.Length} linked cell area(s) for path '{queryPath}'.");

            // --------------------------- 6. Validate cell values ---------------------------
            // Expected values correspond to the <Item> elements (excluding the header cell).
            List<string> expectedValues = new List<string> { "Alpha", "Beta", "Gamma" };
            List<string> actualValues = new List<string>();

            foreach (CellArea area in linkedAreas)
            {
                // Iterate through each cell in the area (normally a single cell).
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        string cellValue = worksheet.Cells[row, col].StringValue;

                        // Skip the header cell that contains the element name ("Item").
                        if (row == 0 && col == 0 && cellValue == "Item")
                            continue;

                        actualValues.Add(cellValue);
                        Console.WriteLine($"Linked cell [{row}, {col}] = \"{cellValue}\"");
                    }
                }
            }

            // Compare expected vs actual.
            bool validationPassed = true;

            if (actualValues.Count != expectedValues.Count)
            {
                Console.WriteLine($"Validation failed: expected {expectedValues.Count} values but found {actualValues.Count}.");
                validationPassed = false;
            }
            else
            {
                for (int i = 0; i < expectedValues.Count; i++)
                {
                    if (!string.Equals(expectedValues[i], actualValues[i], StringComparison.Ordinal))
                    {
                        Console.WriteLine($"Mismatch at index {i}: expected \"{expectedValues[i]}\", got \"{actualValues[i]}\".");
                        validationPassed = false;
                    }
                }
            }

            // --------------------------- 7. Report result ---------------------------
            if (validationPassed)
            {
                Console.WriteLine("All linked cells contain the expected values. Validation succeeded.");
            }
            else
            {
                Console.WriteLine("Validation failed. See messages above for details.");
                return;
            }

            // --------------------------- 8. (Optional) Save workbook for inspection ---------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LinkedCellsValidationResult.xlsx");
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions to prevent the program from crashing.
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
