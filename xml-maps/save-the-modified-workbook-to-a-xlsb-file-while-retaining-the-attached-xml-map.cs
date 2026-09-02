// Title: Save an Excel workbook with an attached XML map to XLSB format using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file containing an XML map, optionally updates a cell, and saves it as a binary .xlsb workbook while keeping the XML map intact with Aspose.Cells. | Show how to configure XlsbSaveOptions in Aspose.Cells to export a workbook to XLSB without losing any attached XML maps. | Provide a robust C# example that verifies the source file, handles exceptions, and writes the workbook to .xlsb preserving all data bindings.
// Common Searches: asp.net how to keep XML map when converting xlsx to xlsb with Aspose.Cells | C# save workbook as xlsb preserving attached XML map | Aspose.Cells XlsbSaveOptions retain xml map during export | convert Excel file to binary format without losing XML map using Aspose | preserve data bindings when saving Excel as .xlsb in C#
// Tags: aspnet aspose.cells save workbook to xlsb with xml map | xlsbsaveoptions retain xml map | c# convert xlsx to xlsb using aspose.cells | binary excel export keep data bindings | aspose.cells workbook save as xlsb preserving mappings

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the input .xlsx file, loads it with Aspose.Cells, optionally modifies a cell, and then saves the workbook as an .xlsb file using XlsbSaveOptions, ensuring that any attached XML map remains intact; errors are caught and reported.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsb";

            // Verify that the source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that already has an XML map attached.
            Workbook workbook = new Workbook(inputPath);

            // Example modification: change a cell value (optional).
            workbook.Worksheets[0].Cells["A1"].PutValue("Modified");

            // Prepare save options for XLSB format.
            // No need to set SaveFormat; passing the options is sufficient.
            XlsbSaveOptions saveOptions = new XlsbSaveOptions();

            // Save the workbook as an XLSB file while preserving the attached XML map.
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
