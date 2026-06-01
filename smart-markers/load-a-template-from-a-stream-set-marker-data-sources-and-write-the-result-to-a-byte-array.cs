using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

public class SmartMarkerProcessor
{
    // Loads an Excel template from a byte array, fills smart markers, and returns the result as a byte array.
    public static byte[] ProcessTemplate(byte[] templateBytes)
    {
        try
        {
            // Load the workbook from the provided byte array
            using (MemoryStream templateStream = new MemoryStream(templateBytes))
            {
                Workbook workbook = new Workbook(templateStream);
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // -------------------------
                // Prepare sample data source
                // -------------------------
                ArrayList data = new ArrayList();

                // Helper to load image bytes safely
                byte[] LoadImage(string path)
                {
                    return File.Exists(path) ? File.ReadAllBytes(path) : Array.Empty<byte>();
                }

                // First record
                data.Add(new
                {
                    Name = "John Doe",
                    Age = 30,
                    Photo = LoadImage("photo1.jpg")
                });

                // Second record
                data.Add(new
                {
                    Name = "Jane Smith",
                    Age = 28,
                    Photo = LoadImage("photo2.jpg")
                });

                // Set the data source for the smart markers (use the name defined in the template)
                designer.SetDataSource("People", data);

                // Process the smart markers
                designer.Process();

                // Save the processed workbook to a memory stream (Excel 97-2003 format)
                using (MemoryStream resultStream = new MemoryStream())
                {
                    workbook.Save(resultStream, SaveFormat.Excel97To2003);
                    return resultStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing template: {ex.Message}");
            return Array.Empty<byte>();
        }
    }

    // Example usage
    public static void Main()
    {
        try
        {
            const string templatePath = "TemplateWithSmartMarkers.xlsx";

            if (!File.Exists(templatePath))
            {
                Console.Error.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load template file into a byte array
            byte[] templateBytes = File.ReadAllBytes(templatePath);

            // Process the template
            byte[] resultBytes = ProcessTemplate(templateBytes);

            if (resultBytes.Length == 0)
            {
                Console.Error.WriteLine("Processing failed; no output generated.");
                return;
            }

            // Write the result to a file to verify
            const string outputPath = "ProcessedResult.xls";
            File.WriteAllBytes(outputPath, resultBytes);

            Console.WriteLine($"Template processed and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}