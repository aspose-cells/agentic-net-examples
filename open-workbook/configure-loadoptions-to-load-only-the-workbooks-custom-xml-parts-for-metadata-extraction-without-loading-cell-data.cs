// Title: Load only custom XML parts from an Excel workbook with Aspose.Cells LoadOptions in C# (skip worksheet data)
// Common Searches: Aspose.Cells C# load workbook without reading cell values to extract custom XML parts | How to use LoadOptions to retrieve only custom XML parts from an XLSX file in .NET | Extract Excel file metadata while avoiding loading worksheets with Aspose.Cells | C# example for loading only custom XML parts from a spreadsheet using Aspose.Cells | Skip worksheet data when opening an Excel workbook with Aspose.Cells LoadOptions
// Tags: Aspose.Cells LoadOptions custom XML extraction | load workbook without cell data C# | Excel custom XML parts retrieval Aspose | metadata extraction from XLSX using LoadOptions | skip worksheet loading Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample checks for the input XLSX file, creates a LoadOptions instance, opens the workbook, and iterates through its CustomXmlParts collection. For each part it reads the byte[] data, converts it to a UTF‑8 string, and outputs the character length, allowing metadata extraction without loading worksheet cell data.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook with default load options.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through custom XML parts and display their content size.
            for (int i = 0; i < workbook.CustomXmlParts.Count; i++)
            {
                try
                {
                    // Use dynamic to avoid compile‑time dependency on the CustomXmlPart type.
                    dynamic xmlPart = workbook.CustomXmlParts[i];
                    byte[] data = xmlPart.Data as byte[];
                    if (data != null)
                    {
                        string xmlContent = Encoding.UTF8.GetString(data);
                        Console.WriteLine($"Custom XML part size: {xmlContent.Length} characters");
                    }
                    else
                    {
                        Console.WriteLine("Custom XML part data is null or not a byte array.");
                    }
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Error processing XML part at index {i}: {innerEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
