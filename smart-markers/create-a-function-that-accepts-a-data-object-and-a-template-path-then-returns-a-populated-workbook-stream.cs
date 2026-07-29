// Title: C# method to fill an Aspose.Cells Excel template with smart markers and return a MemoryStream
// Description: Demonstrates a reusable C# function that validates an Excel template file, loads it with Aspose.Cells, binds any object to a smart‑marker (e.g., "Data") using WorkbookDesigner, processes the markers, and saves the populated workbook as an XLSX MemoryStream ready for web responses or further processing.
// Keywords: Aspose.Cells | C# smart markers | populate Excel template | WorkbookDesigner SetDataSource | MemoryStream Excel output | load template from path | save workbook to stream | Excel report generation | dynamic Excel export
// Common Searches: how to bind data to smart markers Aspose.Cells C# | populate Excel template and get MemoryStream | Aspose.Cells return workbook as stream | C# function to process smart markers in Excel | generate Excel file from template without saving to disk
// Developer Intent: Create a reusable routine that fills an Aspose.Cells Excel template with smart‑marker data and returns the result as a MemoryStream.
// Use Cases: Generate on‑the‑fly reports in an ASP.NET Core API by streaming the populated workbook directly to the client. | Produce downloadable invoices or receipts where order data is bound to a template and sent as an XLSX file. | Batch‑process multiple templates in a background service, converting each populated workbook to a stream for storage or email attachment.
// AI Prompts: Write a C# method that accepts a list of POCO objects, binds it to a smart marker named "Items" in an Aspose.Cells template, and returns the workbook as a MemoryStream. | Add comprehensive error handling to the PopulateWorkbook function to check for missing smart markers, null data sources, and unsupported file formats. | Show how to call PopulateWorkbook from an ASP.NET Core controller and return the generated Excel file using FileStreamResult.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates a reusable C# function that validates an Excel template file, loads it with Aspose.Cells, binds any object to a smart‑marker (e.g., "Data") using WorkbookDesigner, processes the markers, and saves the populated workbook as an XLSX MemoryStream ready for web responses or further processing.
public static class WorkbookHelper
{
    /// <param name="data">The data source to bind to the smart markers in the template.</param>
    /// <param name="templatePath">Full path to the Excel template file containing smart markers.</param>
    /// <returns>A MemoryStream containing the populated workbook in XLSX format.</returns>
    public static MemoryStream PopulateWorkbook(object data, string templatePath)
    {
        // Verify that the template file exists to avoid FileNotFoundException.
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"Template file not found: {templatePath}");

        // Load the template workbook.
        Workbook workbook = new Workbook(templatePath);

        // Initialize the WorkbookDesigner with the loaded workbook.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Bind the provided data object to a smart marker name (e.g., "Data").
        designer.SetDataSource("Data", data);

        // Process the smart markers and fill the workbook with data.
        designer.Process();

        // Save the populated workbook into a memory stream.
        MemoryStream resultStream = new MemoryStream();
        workbook.Save(resultStream, SaveFormat.Xlsx);
        resultStream.Position = 0; // Reset stream position for reading.

        return resultStream;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Expecting two arguments: template path and output path.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <templatePath> <outputPath>");
                return;
            }

            string templatePath = args[0];
            string outputPath = args[1];

            // Ensure the template file exists.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Example data object; replace with actual data as needed.
            var data = new
            {
                Name = "John Doe",
                Age = 30
            };

            // Populate the workbook using the helper.
            using (MemoryStream ms = WorkbookHelper.PopulateWorkbook(data, templatePath))
            {
                // Write the resulting workbook to the specified output file.
                File.WriteAllBytes(outputPath, ms.ToArray());
                Console.WriteLine($"Workbook successfully saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
