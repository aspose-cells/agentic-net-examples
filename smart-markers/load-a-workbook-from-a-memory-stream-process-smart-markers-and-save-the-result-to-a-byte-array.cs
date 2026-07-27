using System;
using System.Data;
using System.IO;
using Aspose.Cells;

public class SmartMarkerProcessor
{
    /// <summary>
    /// Loads a workbook from a memory stream, processes smart markers, and returns the result as a byte array.
    /// </summary>
    /// <param name="templateBytes">The Excel template containing smart markers.</param>
    /// <returns>Byte array of the processed workbook.</returns>
    public static byte[] ProcessSmartMarkers(byte[] templateBytes)
    {
        // Load the workbook from the provided byte array using the Workbook(Stream) constructor.
        using (MemoryStream inputStream = new MemoryStream(templateBytes))
        {
            Workbook workbook = new Workbook(inputStream);

            // Create a WorkbookDesigner for smart marker processing.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // ----- Sample data source -------------------------------------------------
            // In a real scenario replace this with your actual data source.
            DataTable dt = new DataTable("SampleData");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Value", typeof(double));
            dt.Rows.Add("Item A", 10.5);
            dt.Rows.Add("Item B", 20.0);
            // -------------------------------------------------------------------------

            // Bind the data source to the designer.
            designer.SetDataSource(dt);

            // Process all smart markers in the workbook.
            designer.Process();

            // Save the processed workbook to a memory stream using the provided SaveToStream method.
            using (MemoryStream outputStream = workbook.SaveToStream())
            {
                // Return the underlying byte array.
                return outputStream.ToArray();
            }
        }
    }

    // Example usage
    public static void Main()
    {
        // Assume we have an existing Excel template with smart markers as a byte array.
        // For demonstration, create a simple workbook with a smart marker.
        Workbook template = new Workbook();
        Worksheet ws = template.Worksheets[0];
        ws.Cells["A1"].PutValue("&=SampleData.Name");
        ws.Cells["B1"].PutValue("&=SampleData.Value");

        // Save the template to a byte array.
        byte[] templateBytes;
        using (MemoryStream ms = new MemoryStream())
        {
            template.Save(ms, SaveFormat.Xlsx);
            templateBytes = ms.ToArray();
        }

        // Process the smart markers.
        byte[] resultBytes = ProcessSmartMarkers(templateBytes);

        // Optionally, write the result to a file to verify.
        File.WriteAllBytes("ProcessedResult.xlsx", resultBytes);
        Console.WriteLine("Smart markers processed and result saved to ProcessedResult.xlsx");
    }
}