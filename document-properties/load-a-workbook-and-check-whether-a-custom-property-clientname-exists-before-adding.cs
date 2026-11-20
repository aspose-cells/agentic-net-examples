// Title: Load an Excel workbook and add a 'ClientName' custom document property only if it does not already exist using Aspose.Cells for .NET
// Common Searches: how to add a custom document property in Aspose.Cells only when it is missing | C# check if a custom property exists before adding with Aspose.Cells | prevent duplicate custom document properties in an Excel file using Aspose.Cells | conditionally set a workbook custom property in .NET | Aspose.Cells verify custom property presence before insertion
// Tags: conditional custom document property addition Aspose.Cells | check custom property existence .NET | add custom property to Excel workbook Aspose.Cells | avoid duplicate custom properties Excel | load workbook modify custom properties Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads 'input.xlsx' with Aspose.Cells, checks the workbook's CustomDocumentProperties collection for a property named 'ClientName', adds it with the value 'Acme Corporation' only if it is not already present, and saves the result to 'output.xlsx'. It includes file existence validation and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of custom document properties
            var customProps = workbook.CustomDocumentProperties;

            // Add the custom property if it does not already exist
            if (!customProps.Contains("ClientName"))
            {
                customProps.Add("ClientName", "Acme Corporation");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
