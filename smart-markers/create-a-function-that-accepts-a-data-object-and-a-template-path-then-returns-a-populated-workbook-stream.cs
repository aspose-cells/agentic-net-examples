// Title: Populate an Excel template with smart markers and return a MemoryStream using Aspose.Cells (C#)
// Description: Shows how to load an Excel template, bind a .NET object to its smart markers with WorkbookDesigner, process the markers, and stream the resulting workbook (XLS/XLSX) without creating intermediate files.
// Keywords: Aspose.Cells | WorkbookDesigner | smart markers | populate Excel template | C# MemoryStream | Excel report generation | template data binding | stream Excel output | XLSX export | cloud service Excel
// Common Searches: Aspose.Cells populate template from object C# | WorkbookDesigner SetDataSource example | return Excel file as MemoryStream Aspose | process smart markers without saving file | generate Excel report from JSON using Aspose.Cells
// Developer Intent: Create a reusable method that loads an Excel template, binds a supplied data object to its smart markers, processes the markers, and returns the filled workbook as a MemoryStream.
// Use Cases: Serve personalized Excel reports directly from a web API by streaming the workbook to the HTTP response. | Automate invoice generation where order data is bound to a template and the result is attached to an email as a byte array. | Build a serverless function that receives JSON payloads, fills a predefined Excel template, and returns the file stream for downstream processing.
// AI Prompts: Write a C# function that accepts a list of objects and an .xlsx template, uses Aspose.Cells WorkbookDesigner to fill smart markers for each item, and returns a single MemoryStream with all rows merged. | Enhance the PopulateWorkbook method with detailed error handling for missing template files, absent smart markers, and unsupported data types, returning clear exception messages. | Modify the example to output the workbook in XLSX format instead of XLS while still returning a MemoryStream, and show how to set the appropriate save options.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load an Excel template, bind a .NET object to its smart markers with WorkbookDesigner, process the markers, and stream the resulting workbook (XLS/XLSX) without creating intermediate files.
public static class WorkbookHelper
{
    /// <param name="data">The data source object to bind to smart markers in the template.</param>
    /// <param name="templatePath">Full path to the Excel template file containing smart markers.</param>
    /// <returns>A MemoryStream containing the populated workbook (XLS format).</returns>
    public static MemoryStream PopulateWorkbook(object data, string templatePath)
    {
        // Ensure the template file exists to avoid FileNotFoundException.
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"Template file not found: {templatePath}");

        // Load the template workbook.
        Workbook workbook = new Workbook(templatePath);

        // Associate the workbook with a designer for smart marker processing.
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Bind the data source to a marker name (e.g., "Data").
        designer.SetDataSource("Data", data);

        // Process smart markers.
        designer.Process();

        // Return the populated workbook as a memory stream (XLS format).
        return workbook.SaveToStream();
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Define the path to the template file.
            string templatePath = "template.xlsx";

            // Verify the template exists before proceeding.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Example data object matching the smart markers in the template.
            var data = new
            {
                Name = "John Doe",
                Age = 30,
                Email = "john.doe@example.com"
            };

            // Populate the workbook.
            using (MemoryStream resultStream = WorkbookHelper.PopulateWorkbook(data, templatePath))
            {
                // Save the result to a file for verification.
                string outputPath = "output.xls";
                File.WriteAllBytes(outputPath, resultStream.ToArray());
                Console.WriteLine($"Workbook generated successfully: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
