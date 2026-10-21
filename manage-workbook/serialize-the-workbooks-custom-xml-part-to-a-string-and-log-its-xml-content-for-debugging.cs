// Title: Read and log the first custom XML part from an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to get custom xml part data from an Excel file using Aspose.Cells C# | Aspose.Cells read custom XML part and output as string | debug custom xml content in workbook with Aspose.Cells .NET | retrieve first custom XML part from .xlsx using Aspose.Cells | convert custom XML part bytes to UTF8 string Aspose.Cells
// Tags: read custom xml part Aspose.Cells | convert xml part bytes to UTF8 string | log workbook custom xml for debugging | load excel workbook with Aspose.Cells .NET | handle password protected workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example loads an Excel workbook with Aspose.Cells, checks for custom XML parts, converts the first part's binary data to a UTF‑8 string, and writes the XML to the console for debugging, while also handling password‑protected files and general errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (no password supplied). If the file is password‑protected,
            // a CellsException will be thrown and handled in the catch block.
            Workbook workbook = new Workbook(inputPath);

            // Check for custom XML parts
            if (workbook.CustomXmlParts.Count > 0)
            {
                // Serialize the first custom XML part to a string
                var customXmlPart = workbook.CustomXmlParts[0];
                string xmlContent = Encoding.UTF8.GetString(customXmlPart.Data);

                // Output the XML content for debugging purposes
                Console.WriteLine("Custom XML Part Content:");
                Console.WriteLine(xmlContent);
            }
            else
            {
                Console.WriteLine("No custom XML parts found in the workbook.");
            }
        }
        catch (CellsException ex) when (ex.Message.Contains("Please provide password"))
        {
            Console.WriteLine("The workbook is password protected. Please provide a valid password.");
            // Optionally, you could prompt the user for a password and reload with LoadOptions.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
