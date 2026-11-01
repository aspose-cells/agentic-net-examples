// Title: How to set IsNillable = true for a shared custom XML property across multiple Excel workbooks using Aspose.Cells for .NET (API limitation note)
// Common Searches: how to set IsNillable true for a custom XML element in Aspose.Cells C# | make a shared custom XML property optional in multiple Excel files with Aspose.Cells | Aspose.Cells .NET custom XML part nillable attribute not exposed | workaround to add nillable attribute to Excel custom XML using Aspose.Cells | set optional metadata for shared custom XML part across workbooks in C#
// Tags: custom XML part IsNillable Aspose.Cells | optional metadata Excel custom XML C# | shared custom XML property across workbooks | modify custom XML parts Aspose.Cells .NET | nillable attribute handling in Excel with Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, iterates through its CustomXmlParts (noting that Aspose.Cells does not expose an IsNillable property), provides a placeholder for any required XML manipulation, and then saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all custom XML parts.
                // Note: Aspose.Cells CustomXmlPart does not expose an IsNillable property.
                // If needed, additional processing of the XML data can be performed here.
                foreach (var xmlPart in workbook.CustomXmlParts)
                {
                    // Placeholder for any required custom XML manipulation.
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
