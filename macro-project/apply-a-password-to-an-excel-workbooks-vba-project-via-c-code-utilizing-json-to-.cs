using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaPasswordDemo
{
    // Model for JSON credentials
    public class Credentials
    {
        public string Password { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load JSON file that contains the VBA project password
            // Example JSON content: { "Password": "mySecret123" }
            string jsonPath = "credentials.json";
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"JSON file not found: {jsonPath}");
                return;
            }

            string jsonContent = File.ReadAllText(jsonPath);
            Credentials creds = JsonSerializer.Deserialize<Credentials>(jsonContent, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (creds == null || string.IsNullOrEmpty(creds.Password))
            {
                Console.WriteLine("Password not found in JSON.");
                return;
            }

            // Create a new workbook (macro-enabled format will be used on save)
            Workbook workbook = new Workbook();

            // Protect the VBA project with the password from JSON
            // 'true' locks the project for viewing as well
            workbook.VbaProject.Protect(true, creds.Password);

            // Save the workbook as a macro-enabled file
            string outputPath = "ProtectedVbaProject.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved with VBA project protected. Path: {outputPath}");
        }
    }
}