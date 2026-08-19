// Title: C# – Load Excel template from a Stream, apply Smart Markers with a DataTable, and get the result as a byte array (Aspose.Cells)
// Description: Demonstrates how to read an Excel workbook template from a Stream, bind a DataTable as the Smart Marker data source using WorkbookDesigner, process the markers, and return the populated workbook as a byte array via a MemoryStream.
// Keywords: Aspose.Cells | .NET | C# | WorkbookDesigner | Smart Markers | load workbook from stream | DataTable data source | process smart markers | save to MemoryStream | byte array output | in‑memory Excel file | API response Excel | email attachment Excel
// Common Searches: Aspose.Cells load template from stream C# | How to use Smart Markers with DataTable in Aspose.Cells | Convert processed workbook to byte array .NET | WorkbookDesigner process smart markers from stream | Return Excel file as byte[] in ASP.NET Core
// Developer Intent: Read an Excel template from a Stream, fill its Smart Markers using a DataTable, and obtain the final workbook as a byte array.
// Use Cases: Generate product catalogs on‑the‑fly by streaming a template, populating Smart Markers with database data, and sending the byte array as an email attachment. | Expose a REST endpoint that accepts an uploaded Excel template stream, processes Smart Markers, and returns the resulting file as a byte[] in the HTTP response. | Batch‑process multiple template streams in a background service, each with its own DataTable, and store the resulting byte arrays in a document database.
// AI Prompts: Write C# code that reads an Excel template from a Stream, sets a DataTable as the Smart Marker source with WorkbookDesigner, processes the markers, and returns the workbook as a byte array. | Explain performance best practices for handling large Excel template streams when using Aspose.Cells Smart Markers and outputting a byte array. | Show how to replace the DataTable source with a List<T> collection in the Smart Marker example.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Demonstrates how to read an Excel workbook template from a Stream, bind a DataTable as the Smart Marker data source using WorkbookDesigner, process the markers, and return the populated workbook as a byte array via a MemoryStream.
public class SmartMarkerProcessor
{
    // Loads a workbook template from a stream, fills smart markers, and returns the result as a byte array.
    public static byte[] ProcessTemplate(Stream templateStream)
    {
        // Load the workbook from the provided stream.
        Workbook workbook = new Workbook(templateStream);

        // Initialize the WorkbookDesigner with the loaded workbook.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // ----- Prepare sample data source -----
        // Example using a DataTable as a data source.
        DataTable table = new DataTable("Products");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Price", typeof(double));

        table.Rows.Add("Apple", 1.20);
        table.Rows.Add("Banana", 0.80);
        table.Rows.Add("Cherry", 2.50);
        // -------------------------------------

        // Set the data source for the smart markers.
        designer.SetDataSource(table);

        // Process the smart markers and populate the worksheet.
        designer.Process();

        // Save the processed workbook to a memory stream.
        using (MemoryStream resultStream = workbook.SaveToStream())
        {
            // Convert the memory stream to a byte array.
            return resultStream.ToArray();
        }
    }

    // Example usage.
    public static void Run()
    {
        const string templatePath = "Template.xlsx";
        const string resultPath = "Result.xlsx";

        try
        {
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Open the template file safely.
            using (FileStream fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
            {
                byte[] output = ProcessTemplate(fs);
                File.WriteAllBytes(resultPath, output);
                Console.WriteLine($"Template processed and saved to {resultPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Entry point.
    public static void Main(string[] args)
    {
        Run();
    }
}
