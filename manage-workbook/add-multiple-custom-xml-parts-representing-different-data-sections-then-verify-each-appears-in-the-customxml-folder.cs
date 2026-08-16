// Title: Add Multiple Custom XML Parts to an Excel Workbook and Verify Them with Aspose.Cells for .NET
// Description: Shows how to create a Workbook, embed three distinct custom XML parts (each representing a separate data section), assign unique GUID IDs, save the file, reload it, and confirm that all parts appear in the workbook's customXml folder.
// Keywords: Aspose.Cells | custom XML parts | multiple custom XML parts | C# | .NET | Excel workbook | customXml folder | save and reload workbook | verify custom XML | GUID ID | embed XML in Excel | CustomXmlParts collection
// Common Searches: how to add several custom XML parts to an Excel file using Aspose.Cells | verify custom XML parts count after saving a workbook | retrieve custom XML part IDs from a loaded workbook C# | Aspose.Cells example for customXml folder | embed multiple XML sections in Excel with Aspose.Cells
// Developer Intent: Embed several custom XML sections in an Excel file and ensure they persist after saving.
// Use Cases: Store section‑specific metadata (e.g., configuration, lookup tables) as separate XML parts for downstream processing. | Maintain independent XML payloads that can be updated without modifying worksheet data. | Validate that embedded XML survives the save/load cycle to meet document‑exchange standards.
// AI Prompts: Generate C# code using Aspose.Cells to add a list of XML strings as separate custom XML parts, assign GUID IDs, and save the workbook. | Show how to iterate through CustomXmlParts in a loaded workbook and output each part's ID and XML content. | Explain how to map custom XML parts to specific worksheet sections for data binding with Aspose.Cells.

using System;
using System.Text;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsExamples
{
    // Shows how to create a Workbook, embed three distinct custom XML parts (each representing a separate data section), assign unique GUID IDs, save the file, reload it, and confirm that all parts appear in the workbook's customXml folder.
    public class MultipleCustomXmlPartsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Prepare first custom XML part data
                string xmlData1 = "<Section1><Item>Value1</Item></Section1>";
                byte[] xmlBytes1 = Encoding.UTF8.GetBytes(xmlData1);

                // Prepare second custom XML part data
                string xmlData2 = "<Section2><Item>Value2</Item></Section2>";
                byte[] xmlBytes2 = Encoding.UTF8.GetBytes(xmlData2);

                // Prepare third custom XML part data
                string xmlData3 = "<Section3><Item>Value3</Item></Section3>";
                byte[] xmlBytes3 = Encoding.UTF8.GetBytes(xmlData3);

                // Add the custom XML parts to the workbook (no schema is provided)
                int index1 = workbook.CustomXmlParts.Add(xmlBytes1, null);
                int index2 = workbook.CustomXmlParts.Add(xmlBytes2, null);
                int index3 = workbook.CustomXmlParts.Add(xmlBytes3, null);

                // Optionally assign explicit IDs (useful for later verification)
                workbook.CustomXmlParts[index1].ID = Guid.NewGuid().ToString();
                workbook.CustomXmlParts[index2].ID = Guid.NewGuid().ToString();
                workbook.CustomXmlParts[index3].ID = Guid.NewGuid().ToString();

                // Save the workbook to a file
                string filePath = "MultipleCustomXmlParts.xlsx";
                workbook.Save(filePath);

                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Reload the workbook from the saved file
                Workbook loadedWorkbook = new Workbook(filePath);

                // Verify that all custom XML parts are present
                int expectedCount = 3;
                int actualCount = loadedWorkbook.CustomXmlParts.Count;
                Console.WriteLine($"Expected custom XML parts count: {expectedCount}");
                Console.WriteLine($"Actual custom XML parts count:   {actualCount}");

                // List each part's ID and XML content to confirm they are stored correctly
                for (int i = 0; i < actualCount; i++)
                {
                    CustomXmlPart part = loadedWorkbook.CustomXmlParts[i];
                    string partId = part.ID;
                    string partXml = Encoding.UTF8.GetString(part.Data);
                    Console.WriteLine($"Part {i + 1} - ID: {partId}");
                    Console.WriteLine($"Part {i + 1} - XML: {partXml}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MultipleCustomXmlPartsDemo.Run();
        }
    }
}
