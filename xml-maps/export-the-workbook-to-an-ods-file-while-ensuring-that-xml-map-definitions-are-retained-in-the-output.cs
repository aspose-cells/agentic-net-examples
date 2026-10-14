// Title: Export an Excel workbook to ODS format while preserving XML map definitions with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook and saves it as an .ods file, keeping any XML map definitions intact using Aspose.Cells. | Show how to configure OdsSaveOptions in Aspose.Cells so that XML maps are retained when converting an Excel workbook to ODS in C#.
// Common Searches: Aspose.Cells keep XML map definitions when converting Excel to ODS in C# | C# ODS export preserving XML maps with Aspose.Cells | How to retain data mapping after saving a workbook as ODS using Aspose.Cells
// Tags: ODS export preserving XML maps Aspose.Cells | C# OdsSaveOptions retain data mapping | Excel to ODS conversion with XML map support | Aspose.Cells workbook save ODS with schema

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing .xlsx file, configures OdsSaveOptions, and saves the workbook as an .ods file while retaining any XML map definitions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.ods";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure ODS save options (XML maps are retained by default if supported)
            OdsSaveOptions odsOptions = new OdsSaveOptions(SaveFormat.Ods);

            // Save the workbook as an ODS file with the specified options
            workbook.Save(outputPath, odsOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
