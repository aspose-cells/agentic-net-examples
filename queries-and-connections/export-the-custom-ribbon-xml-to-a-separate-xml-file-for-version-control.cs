// Title: Export a workbook's custom ribbon XML to a separate .xml file using Aspose.Cells for .NET
// Common Searches: how to extract the custom ribbon XML part from an .xlsx file with Aspose.Cells in C# | saving Excel custom UI XML from a workbook to an external file for source control using Aspose.Cells | C# code to read custom XML parts of a workbook and write them to a .xml file with Aspose.Cells | retrieve custom ribbon definition from an Excel file using Aspose.Cells .NET API | export custom UI XML from a spreadsheet for version control in .NET
// Tags: export custom ribbon xml Aspose.Cells | read workbook custom xml parts C# | write xml file utf-8 Aspose.Cells | store excel custom ui definition in source control | extract custom UI part from .xlsx using .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads an Excel workbook, checks for custom XML parts, extracts the first part containing the ribbon UI definition, and writes it as a UTF‑8 .xml file, handling missing files and exceptions.
class ExportRibbonXml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "RibbonCustomUI.xml";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains the custom ribbon XML.
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook has at least one custom XML part.
            if (workbook.CustomXmlParts.Count > 0)
            {
                // Retrieve the XML content of the first custom part.
                // CustomXmlPart.Data returns a byte array; convert it to a UTF‑8 string.
                byte[] dataBytes = workbook.CustomXmlParts[0].Data;
                string ribbonXml = Encoding.UTF8.GetString(dataBytes);

                // Export the ribbon XML to a separate file for version control.
                File.WriteAllText(outputPath, ribbonXml, Encoding.UTF8);

                Console.WriteLine($"Ribbon XML exported successfully to \"{outputPath}\"");
            }
            else
            {
                Console.WriteLine("No custom XML parts found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
