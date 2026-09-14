// Title: Process smart markers in an Excel workbook loaded from a byte array and return the result as a byte array (Excel97‑2003) using Aspose.Cells for .NET
// AI Prompts: Write C# code that takes a byte[] with an Excel template, binds a DataTable to Aspose.Cells WorkbookDesigner, processes all smart markers, and returns the modified workbook as a byte[] in Excel97‑2003 format. | Show how to load an Excel file from a MemoryStream, apply smart markers with Aspose.Cells, and save the processed workbook to another MemoryStream without touching the file system.
// Common Searches: asp.net load excel from byte array and process smart markers | asp.net core convert processed workbook to byte array using Aspose.Cells | how to bind a DataTable to WorkbookDesigner for smart markers in memory | save Aspose.Cells workbook as Excel97-2003 format without creating a file | process smart markers from template.xlsx stored in a memory stream c#
// Tags: WorkbookDesigner smart markers binding | load workbook from byte array Aspose.Cells | save workbook as Excel97-2003 byte array | in-memory Excel processing Aspose.Cells | DataTable data source for smart markers

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook from a byte array via a MemoryStream, binding a DataTable to WorkbookDesigner to resolve smart markers, and saving the resulting workbook to another MemoryStream in Excel97‑2003 format, returning the final file as a byte array.
public static class SmartMarkerProcessor
{
    /// <param name="templateBytes">Byte array containing the source workbook with smart markers.</param>
    /// <returns>Byte array of the processed workbook.</returns>
    public static byte[] ProcessSmartMarkers(byte[] templateBytes)
    {
        // Load the workbook from the input memory stream
        using (MemoryStream inputStream = new MemoryStream(templateBytes))
        {
            Workbook workbook = new Workbook(inputStream);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // ----- Sample data source for demonstration -----
            // Create a DataTable that matches the smart markers in the template
            DataTable data = new DataTable("SampleData");
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Value", typeof(double));

            // Populate the table with sample rows
            data.Rows.Add("Item1", 10.5);
            data.Rows.Add("Item2", 20.75);
            data.Rows.Add("Item3", 30.0);

            // Bind the data source to the designer
            designer.SetDataSource(data);
            // ------------------------------------------------

            // Process all smart markers in the workbook
            designer.Process();

            // Save the processed workbook to a memory stream (Excel97-2003 format)
            using (MemoryStream outputStream = new MemoryStream())
            {
                workbook.Save(outputStream, SaveFormat.Excel97To2003);
                return outputStream.ToArray();
            }
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string templatePath = args.Length > 0 ? args[0] : "template.xlsx";

            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            byte[] templateBytes = File.ReadAllBytes(templatePath);
            byte[] resultBytes = SmartMarkerProcessor.ProcessSmartMarkers(templateBytes);

            string outputPath = "output.xls";
            File.WriteAllBytes(outputPath, resultBytes);
            Console.WriteLine($"Processed workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
