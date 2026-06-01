using System;
using System.IO;
using Aspose.Cells;

public class SmartMarkerProcessor
{
    // Entry point required for console application
    public static void Main(string[] args)
    {
        try
        {
            // Path to the template workbook
            const string templatePath = "template.xlsx";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load template bytes
            byte[] templateBytes = File.ReadAllBytes(templatePath);

            // Process the workbook and obtain the result bytes
            byte[] resultBytes = ProcessWorkbook(templateBytes);

            // Save the processed workbook
            const string outputPath = "result.xls";
            File.WriteAllBytes(outputPath, resultBytes);
            Console.WriteLine($"Processed workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // This method demonstrates loading a workbook from a memory stream,
    // processing smart markers, and returning the result as a byte array.
    public static byte[] ProcessWorkbook(byte[] templateBytes)
    {
        try
        {
            // Load the workbook from the provided byte array using a MemoryStream.
            using (MemoryStream inputStream = new MemoryStream(templateBytes))
            {
                // Initialize the Workbook from the stream.
                Workbook workbook = new Workbook(inputStream);

                // Create a WorkbookDesigner and associate it with the loaded workbook.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Example data source: a simple JSON string.
                // Adjust the JSON and data source name to match the smart markers in your template.
                string jsonData = @"{""Name"":""John Doe"",""Age"":30}";
                designer.SetJsonDataSource("Employee", jsonData);

                // Process all smart markers in the workbook.
                designer.Process();

                // Save the processed workbook to a MemoryStream (Excel97-2003 format) and obtain the byte array.
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Excel97To2003);
                    return outputStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            // Propagate the exception after logging
            Console.WriteLine($"Processing error: {ex.Message}");
            throw;
        }
    }
}