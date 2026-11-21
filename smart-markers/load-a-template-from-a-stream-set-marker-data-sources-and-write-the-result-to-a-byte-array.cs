// Title: C# – Load an Excel template from a Stream, bind a DataTable to smart markers with WorkbookDesigner, and return the processed workbook as a byte array (XLSX) using Aspose.Cells
// Common Searches: how to bind a DataTable to smart markers in C# using Aspose.Cells WorkbookDesigner | process an Excel smart marker template from a Stream and get a byte[] output in .NET | save a processed smart marker workbook to a memory stream as XLSX with Aspose.Cells | load Excel template from FileStream and generate byte array result using smart markers | Aspose.Cells example for converting a smart marker template to a byte array in C#
// Tags: WorkbookDesigner SetDataSource DataTable smart markers | load Excel template from stream Aspose.Cells | export processed workbook to byte array XLSX | process smart markers in memory using C# | Aspose.Cells generate byte[] from smart marker template

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The ProcessTemplate method reads an Excel template workbook from a Stream, uses WorkbookDesigner to bind a DataTable named "DataSource" to smart markers, processes the markers, and then saves the resulting workbook to a MemoryStream in XLSX format, returning the content as a byte array. Errors are wrapped in an ApplicationException for clearer reporting.
public class SmartMarkerExample
{
    /// <param name="templateStream">Stream of the template workbook.</param>
    /// <param name="dataSource">DataTable that supplies data for the smart markers.</param>
    /// <returns>Byte array of the processed workbook (XLSX format).</returns>
    public byte[] ProcessTemplate(Stream templateStream, DataTable dataSource)
    {
        if (templateStream == null) throw new ArgumentNullException(nameof(templateStream));
        if (dataSource == null) throw new ArgumentNullException(nameof(dataSource));

        try
        {
            // Load the workbook template from the provided stream
            Workbook workbook = new Workbook(templateStream);

            // Use WorkbookDesigner to bind data and process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("DataSource", dataSource);
            designer.Process();

            // Save the processed workbook to a memory stream in XLSX format
            using (MemoryStream outputStream = new MemoryStream())
            {
                workbook.Save(outputStream, SaveFormat.Xlsx);
                return outputStream.ToArray();
            }
        }
        catch (Exception ex)
        {
            // Wrap any exception for clearer error reporting
            throw new ApplicationException("Error processing smart marker template.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Expect the template file path as the first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: SmartMarkerExample <templateFilePath>");
                return;
            }

            string templatePath = args[0];

            // Prevent FileNotFoundException by checking existence first
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Create a sample DataTable as data source
            DataTable dt = new DataTable("DataSource");
            dt.Columns.Add("ColumnName", typeof(string));
            dt.Rows.Add("Value 1");
            dt.Rows.Add("Value 2");
            dt.Rows.Add("Value 3");

            // Process the template
            using (FileStream fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
            {
                SmartMarkerExample example = new SmartMarkerExample();
                byte[] result = example.ProcessTemplate(fs, dt);

                // Write the result to an output file in the same directory
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(templatePath) ?? string.Empty,
                    Path.GetFileNameWithoutExtension(templatePath) + "_Processed.xlsx");

                File.WriteAllBytes(outputPath, result);
                Console.WriteLine($"Processed file saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
