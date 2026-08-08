// Title: Process Smart Markers from a MemoryStream and Return a Byte Array – Aspose.Cells for .NET
// Description: Loads an Excel template from a byte array using a MemoryStream, binds a DataTable to WorkbookDesigner, processes all smart markers, and saves the populated workbook back to a MemoryStream. The method returns the final workbook as a byte array, enabling in‑memory Excel generation without touching the file system.
// Keywords: Aspose.Cells | WorkbookDesigner | smart markers | memory stream | byte array | C# | .NET | in‑memory Excel | template processing | DataTable binding | Excel report generation
// Common Searches: Aspose.Cells load workbook from byte array | process smart markers from memory stream C# | save processed workbook to byte array Aspose | WorkbookDesigner example without file system | generate Excel from template stored in database
// Developer Intent: Load a workbook from a byte array, apply smart‑marker data, and obtain the result as a byte array.
// Use Cases: Web API that receives an Excel template as a BLOB, fills smart markers, and returns the file as a downloadable byte stream. | Background service that creates dynamic Excel reports in memory and uploads them directly to cloud storage. | Desktop application that stores smart‑marker templates in a database, processes them on demand, and saves the output without creating temporary files.
// AI Prompts: Write a C# method that accepts a template byte array, binds a DataTable to WorkbookDesigner, processes smart markers, and returns the processed workbook as a byte array using Aspose.Cells. | Show how to extend the example to handle multiple DataTables for different smart‑marker groups within the same workbook. | Provide code to stream the resulting byte array from this method directly to an ASP.NET Core controller action for file download.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Loads an Excel template from a byte array using a MemoryStream, binds a DataTable to WorkbookDesigner, processes all smart markers, and saves the populated workbook back to a MemoryStream. The method returns the final workbook as a byte array, enabling in‑memory Excel generation without touching the file system.
public class SmartMarkerProcessor
{
    // Processes a workbook containing smart markers from a memory stream
    // and returns the resulting workbook as a byte array.
    public static byte[] ProcessSmartMarkers(byte[] templateBytes)
    {
        // Load the template workbook from the provided byte array (memory stream)
        using (MemoryStream templateStream = new MemoryStream(templateBytes))
        {
            // Use the Workbook constructor that accepts a Stream (load rule)
            Workbook workbook = new Workbook(templateStream);

            // Initialize WorkbookDesigner with the loaded workbook (create rule)
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare a simple data source for the smart markers
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Value", typeof(double));
            dt.Rows.Add("Item A", 123.45);
            dt.Rows.Add("Item B", 678.90);

            // Bind the data source to the designer
            designer.SetDataSource(dt);

            // Process all smart markers in the workbook (process rule)
            designer.Process();

            // Save the processed workbook to a MemoryStream (save rule)
            MemoryStream resultStream = workbook.SaveToStream();

            // Return the workbook content as a byte array
            return resultStream.ToArray();
        }
    }

    // Example usage
    public static void Main()
    {
        // ----- Create a sample template workbook with smart markers -----
        Workbook template = new Workbook();
        Worksheet sheet = template.Worksheets[0];
        // Smart markers using the table name "Data"
        sheet.Cells["A1"].PutValue("&=Data.Name");
        sheet.Cells["B1"].PutValue("&=Data.Value");

        // Save the template to a byte array
        using (MemoryStream tmplStream = new MemoryStream())
        {
            template.Save(tmplStream, SaveFormat.Xlsx);
            byte[] templateBytes = tmplStream.ToArray();

            // ----- Process the template -----
            byte[] resultBytes = ProcessSmartMarkers(templateBytes);

            // Optionally write the result to a file to verify
            File.WriteAllBytes("ProcessedResult.xlsx", resultBytes);
            Console.WriteLine("Processed workbook saved as 'ProcessedResult.xlsx'.");
        }
    }
}
