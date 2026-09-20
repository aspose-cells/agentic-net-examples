// Title: Export an Excel workbook that contains XML map definitions to ODS with Aspose.Cells for .NET (XML maps not retained)
// AI Prompts: Write C# code that loads an .xlsx workbook with XML maps and saves it as an .ods file using Aspose.Cells, including error handling for missing input files. | Show how to configure OdsSaveOptions in Aspose.Cells to export a workbook to ODS and explain why XML map definitions are omitted in the resulting file. | Provide a brief explanation of the ExportXmlMap limitation for ODS format in Aspose.Cells and suggest alternative approaches for preserving XML data.
// Common Searches: Aspose.Cells C# export workbook with XML maps to ODS format | why XML map definitions are lost when saving to ODS using Aspose.Cells | how to configure OdsSaveOptions for XLSX to ODS conversion in .NET | C# convert Excel file containing XML maps to ODS with Aspose.Cells | workaround for preserving XML map data when exporting to ODS in Aspose.Cells
// Tags: aspocells ods saveoptions xmlmap limitation | export workbook to ods using aspocells | xml map definitions lost in ods conversion | c# aspocells convert xlsx to ods | workaround preserving xml data in ods export

using Aspose.Cells;
using System;
using System.IO;

namespace ExportToOdsWithXmlMap
{
    // The example loads an existing Excel workbook (input.xlsx) that includes XML map definitions, creates OdsSaveOptions for the ODS format, and saves the workbook as output.ods. Aspose.Cells does not support preserving XML maps when exporting to ODS, so the XML map definitions are not retained in the resulting file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.ods";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains XML map definitions
                Workbook workbook = new Workbook(inputPath);

                // Configure ODS save options (ExportXmlMap is not supported for ODS)
                OdsSaveOptions odsOptions = new OdsSaveOptions(SaveFormat.Ods);

                // Save the workbook as ODS
                workbook.Save(outputPath, odsOptions);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
