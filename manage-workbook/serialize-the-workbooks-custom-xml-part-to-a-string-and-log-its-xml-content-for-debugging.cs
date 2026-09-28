// Title: Serialize workbook custom XML parts to a UTF‑8 string and log their content using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, enumerates Workbook.CustomXmlParts, converts each part's Data byte[] to a UTF‑8 string, and prints the part name and XML to the console with proper error handling. | Show how to add file‑existence verification and exception handling while extracting and debugging custom XML parts from an Excel workbook using Aspose.Cells.
// Common Searches: C# Aspose.Cells how to extract custom XML part from an existing XLSX file | log XML content of workbook custom XML parts for debugging Aspose.Cells | convert custom XML part byte array to string in Aspose.Cells .NET | iterate through Workbook.CustomXmlParts and display XML using Aspose.Cells | handle missing Excel file when reading custom XML parts with Aspose.Cells C#
// Tags: Aspose.Cells serialize custom XML part | C# read workbook custom XML parts | debug Excel custom XML with Aspose.Cells | convert custom XML byte array to UTF-8 string | iterate Workbook.CustomXmlParts Aspose

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an existing XLSX file, checks that the file exists, iterates over the workbook's CustomXmlParts collection, converts each part's byte[] Data to a UTF‑8 string, and writes the part name and XML content to the console while handling file‑not‑found and processing exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all custom XML parts in the workbook
            foreach (var customXmlObj in workbook.CustomXmlParts)
            {
                try
                {
                    // Use dynamic to avoid compile‑time dependency on the exact type name
                    dynamic customXml = customXmlObj;

                    // Convert the XML data (byte array) to a UTF‑8 string
                    string xmlContent = Encoding.UTF8.GetString((byte[])customXml.Data);

                    // Log the name of the custom XML part and its XML content
                    Console.WriteLine($"Custom XML Part Name: {customXml.Name}");
                    Console.WriteLine("XML Content:");
                    Console.WriteLine(xmlContent);
                    Console.WriteLine(new string('-', 80));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process a custom XML part: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
