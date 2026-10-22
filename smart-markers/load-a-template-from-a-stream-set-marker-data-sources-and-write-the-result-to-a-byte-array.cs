// Title: Load an Excel template from a Stream, assign smart‑marker data sources, and retrieve the processed workbook as a byte array using Aspose.Cells for .NET
// AI Prompts: Read an Excel template from a Stream, assign each smart‑marker data source with WorkbookDesigner.SetDataSource, execute the marker processing, and return the workbook as a byte[] in XLSX format. | Refactor the method to accept a strongly‑typed collection of data source objects, process the smart markers, and write the resulting workbook directly to a MemoryStream without creating an intermediate file.
// Common Searches: Aspose.Cells how to load an Excel template from a Stream and process smart markers | C# save processed workbook to a byte array using WorkbookDesigner | Set multiple smart marker data sources programmatically with Aspose.Cells .NET | Convert processed Excel workbook to byte[] without writing to disk Aspose.Cells
// Tags: WorkbookDesigner SetDataSource C# | load workbook from stream Aspose.Cells | process smart markers in memory | save workbook to byte array XLSX | Excel template processing Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The ProcessTemplate method loads an Excel workbook from a Stream, uses WorkbookDesigner to assign smart‑marker data sources from a dictionary, processes all markers, and saves the resulting workbook as an XLSX byte array, enabling in‑memory template handling.
    public class ExcelProcessor
    {
        /// <param name="templateStream">Stream containing the Excel template.</param>
        /// <param name="markerDataSources">Dictionary where key is the marker name and value is the data source object.</param>
        /// <returns>Byte array of the processed workbook.</returns>
        public byte[] ProcessTemplate(Stream templateStream, Dictionary<string, object> markerDataSources)
        {
            try
            {
                // Load the workbook from the provided stream
                Workbook workbook = new Workbook(templateStream);

                // Initialize WorkbookDesigner for smart marker processing
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Set each marker's data source
                foreach (KeyValuePair<string, object> entry in markerDataSources)
                {
                    designer.SetDataSource(entry.Key, entry.Value);
                }

                // Process all smart markers in the workbook
                designer.Process();

                // Save the processed workbook to a memory stream
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
            catch (Exception ex)
            {
                // Wrap and rethrow for caller handling
                throw new ApplicationException("Error processing the Excel template.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect the template file path as the first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the Excel template file as an argument.");
                return;
            }

            string templatePath = args[0];

            // Prevent FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            try
            {
                using (FileStream templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                {
                    // Example: empty data sources; populate as needed
                    var dataSources = new Dictionary<string, object>();

                    ExcelProcessor processor = new ExcelProcessor();
                    byte[] processedBytes = processor.ProcessTemplate(templateStream, dataSources);

                    // Save the processed workbook next to the template
                    string outputPath = Path.Combine(Path.GetDirectoryName(templatePath) ?? string.Empty, "ProcessedOutput.xlsx");
                    File.WriteAllBytes(outputPath, processedBytes);
                    Console.WriteLine($"Processed workbook saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}
