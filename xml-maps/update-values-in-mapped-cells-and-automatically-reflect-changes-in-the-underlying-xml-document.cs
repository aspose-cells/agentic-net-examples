// Title: How to modify XML‑mapped cells in a workbook and save the changes back to an XML file using Aspose.Cells for .NET
// AI Prompts: Load an XML file as a workbook with Aspose.Cells, update the values of cells A1, B2, and C3 that are mapped to XML nodes, and save the workbook as a new XML document. | Create a minimal XML workbook when the source file is missing, then edit the mapped cells and export the updated XML using C# and Aspose.Cells. | Use LoadOptions with LoadFormat.Xml to open an XML‑mapped workbook, set new values for specific cells, and persist the modifications to another XML file.
// Common Searches: Aspose.Cells C# update values of cells linked to XML map and export updated XML | How to write changes from Excel cells back to original XML using Aspose.Cells .NET | Save modified XML‑mapped workbook to a new XML file with Aspose.Cells LoadOptions
// Tags: modify XML-mapped cells Aspose.Cells | export workbook to XML C# | LoadOptions LoadFormat.Xml usage | fallback minimal XML workbook creation | synchronize Excel edits with XML source

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an XML file as a workbook via Aspose.Cells, updates the values of cells A1, B2, and C3 that are mapped to XML nodes, and saves the workbook to a new XML file, creating a minimal XML workbook when the source file is absent.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xml";
                string outputPath = "output.xml";

                // Ensure the input XML file exists; create a minimal workbook if it doesn't.
                if (!File.Exists(inputPath))
                {
                    Workbook tempWb = new Workbook();
                    tempWb.Save(inputPath, SaveFormat.Xml);
                }

                // Load the workbook that is linked to an XML data source.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xml);
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Update values in cells mapped to XML nodes.
                cells["A1"].PutValue("UpdatedValue1");
                cells["B2"].PutValue(12345);
                cells["C3"].PutValue(DateTime.Now);

                // Save the workbook back to XML format.
                workbook.Save(outputPath, SaveFormat.Xml);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
