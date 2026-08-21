// Title: Read and Log Custom XML Parts from an Excel Workbook using Aspose.Cells for .NET
// Description: Shows how to open an .xlsx file with Aspose.Cells, enumerate its CustomXmlPartCollection, convert each part's byte array to a UTF‑8 string, and write the XML and part ID to the console for debugging.
// Keywords: Aspose.Cells custom XML part | C# read custom XML from Excel | convert CustomXmlPart to string | debug Excel custom XML | CustomXmlPartCollection Aspose | UTF-8 XML extraction .NET | Aspose.Cells workbook XML debugging
// Common Searches: How to get XML string from CustomXmlPart in Aspose.Cells | C# list custom XML parts in an Excel workbook | Serialize CustomXmlPart data to string Aspose.Cells | Debug custom XML parts in .xlsx using .NET | Extract embedded XML from Excel with Aspose.Cells
// Developer Intent: Extract every custom XML part from a workbook, convert its data to a UTF‑8 string, and output the XML for troubleshooting.
// Use Cases: Verify the structure of embedded XML before applying transformations. | Log XML content to a file or console to diagnose import/export issues. | Ensure custom XML parts are present and contain valid data prior to automated processing.
// AI Prompts: Create a C# method that returns a list of UTF‑8 XML strings from all CustomXmlParts in a workbook, handling null or empty parts gracefully. | Write code to merge all custom XML parts into one XML document, prepend each part with a comment containing its ID, and save the result to a .txt file. | Generate robust error‑handling that records the ID of any CustomXmlPart that fails during UTF‑8 conversion and continues processing the remaining parts.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsExamples
{
    // Shows how to open an .xlsx file with Aspose.Cells, enumerate its CustomXmlPartCollection, convert each part's byte array to a UTF‑8 string, and write the XML and part ID to the console for debugging.
    public class CustomXmlPartDebugDemo
    {
        public static void Run()
        {
            try
            {
                // Specify the workbook path
                string workbookPath = "input.xlsx";

                // Verify that the file exists to avoid FileNotFoundException
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"File not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook wb = new Workbook(workbookPath);

                // Access custom XML parts collection
                CustomXmlPartCollection xmlParts = wb.CustomXmlParts;

                // Inform if there are no custom XML parts
                if (xmlParts.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any custom XML parts.");
                    return;
                }

                // Iterate through each custom XML part and display its content
                for (int i = 0; i < xmlParts.Count; i++)
                {
                    CustomXmlPart part = xmlParts[i];
                    byte[] dataBytes = part.Data;

                    // Guard against null or empty data
                    if (dataBytes == null || dataBytes.Length == 0)
                    {
                        Console.WriteLine($"Custom XML part at index {i} has no data.");
                        continue;
                    }

                    // Convert byte array to UTF-8 string
                    string xmlContent = Encoding.UTF8.GetString(dataBytes);

                    // Output the XML content for debugging
                    Console.WriteLine($"--- Custom XML Part {i} (ID: {part.ID}) ---");
                    Console.WriteLine(xmlContent);
                    Console.WriteLine("--- End of Part ---");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
