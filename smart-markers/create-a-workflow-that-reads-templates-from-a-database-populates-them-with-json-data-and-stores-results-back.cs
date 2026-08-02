using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsApp
{
    public class TemplateProcessor
    {
        // Entry point for processing all templates
        public void ProcessTemplates()
        {
            try
            {
                // 1. Retrieve the template workbook (as a byte array) from the database
                byte[] templateBytes = GetTemplateFromDatabase(templateId: 1);

                // 2. Load the workbook from the byte array using the Workbook constructor
                Workbook workbook = new Workbook(new MemoryStream(templateBytes));

                // 3. Create a WorkbookDesigner and associate it with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // 4. Obtain JSON data that will populate the smart markers
                string jsonData = GetJsonDataFromDatabase(dataId: 1);

                // 5. Set the JSON string as a data source for the smart markers
                designer.SetJsonDataSource("Data", jsonData);

                // 6. Process the smart markers and fill the workbook with data
                designer.Process();

                // 7. Save the processed workbook back to a byte array
                using (MemoryStream ms = new MemoryStream())
                {
                    designer.Workbook.Save(ms, SaveFormat.Xlsx);
                    byte[] resultBytes = ms.ToArray();

                    // 8. Store the resulting workbook back into the database
                    SaveResultToDatabase(templateId: 1, resultBytes);
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                Console.Error.WriteLine($"Error processing templates: {ex.Message}");
                throw;
            }
        }

        // Placeholder method: replace with actual DB retrieval logic
        private byte[] GetTemplateFromDatabase(int templateId)
        {
            const string templatePath = "Template.xlsx";

            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            return File.ReadAllBytes(templatePath);
        }

        // Placeholder method: replace with actual DB retrieval logic for JSON data
        private string GetJsonDataFromDatabase(int dataId)
        {
            // Example JSON payload; in production, fetch from DB or another service
            return "{\"Name\":\"John Doe\",\"Age\":30,\"City\":\"New York\"}";
        }

        // Placeholder method: replace with actual DB update logic
        private void SaveResultToDatabase(int templateId, byte[] resultBytes)
        {
            const string resultPath = "Result.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(resultPath));
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            File.WriteAllBytes(resultPath, resultBytes);
        }
    }

    public static class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                var processor = new TemplateProcessor();
                processor.ProcessTemplates();
                Console.WriteLine("Template processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}