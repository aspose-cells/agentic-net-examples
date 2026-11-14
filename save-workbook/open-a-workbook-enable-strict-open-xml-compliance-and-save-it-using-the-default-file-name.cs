// Title: How to open or create an Excel workbook, enable strict Open XML compliance, and save it with the default file name using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# open existing workbook or create new and enable strict Open XML compliance | save Excel workbook with default file name using Aspose.Cells .NET | set strict OpenXML compliance when saving a workbook with Aspose.Cells | conditional workbook loading based on file existence in C# Aspose.Cells example | default XLSX format save operation with Aspose.Cells for .NET
// Tags: open or create workbook Aspose.Cells C# | enable strict OpenXML compliance Aspose.Cells | save workbook default XLSX Aspose.Cells | conditional workbook loading Aspose.Cells | file existence check C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks whether 'input.xlsx' exists, loads the workbook if present or creates a new one otherwise, optionally enables strict Open XML compliance, and then saves the workbook back to the same path using the default XLSX format.
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be opened or created
            string filePath = "input.xlsx";

            try
            {
                Workbook workbook;

                if (File.Exists(filePath))
                {
                    // Load the existing workbook from the specified file
                    workbook = new Workbook(filePath);
                }
                else
                {
                    // Create a new workbook when the file does not exist
                    workbook = new Workbook();
                }

                // Save the workbook using the default XLSX format
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(filePath)}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
