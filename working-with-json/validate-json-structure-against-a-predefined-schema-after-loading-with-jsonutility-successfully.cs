// Title: Validate JSON syntax and required fields in an Aspose.Cells worksheet cell using C# and System.Text.Json
// Common Searches: how to parse JSON stored in an Aspose.Cells worksheet cell with System.Text.Json in C# | C# code to check required properties Name, Age, Email in JSON string placed in Excel cell | validate email format in JSON data loaded into an Aspose.Cells workbook | detect syntactically invalid JSON in a cell of an Excel file using .NET | manual JSON schema validation without third‑party library in Aspose.Cells
// Tags: System.Text.Json JSON syntax validation C# | Aspose.Cells write JSON to Excel cell | custom required property validation JSON C# | email format verification in JSON C# | JSON validation without external library .NET

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The program creates an Aspose.Cells workbook, writes a JSON string into cell A1, parses it with System.Text.Json to confirm syntactic correctness, then manually checks that the required properties Name, Age, and Email are present with the correct JSON types, and outputs whether the JSON conforms to the predefined schema.
class Program
{
    static void Main()
    {
        try
        {
            // Sample JSON data to be validated
            string jsonData = @"{
                ""Name"": ""John Doe"",
                ""Age"": 30,
                ""Email"": ""john.doe@example.com""
            }";

            // Predefined JSON schema describing the expected structure
            string jsonSchema = @"{
                ""type"": ""object"",
                ""properties"": {
                    ""Name"": { ""type"": ""string"" },
                    ""Age"": { ""type"": ""integer"" },
                    ""Email"": { ""type"": ""string"", ""format"": ""email"" }
                },
                ""required"": [""Name"", ""Age"", ""Email""]
            }";

            // Create a new workbook and place the JSON string into the first cell
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(jsonData);

            // Validate JSON syntax using System.Text.Json
            bool isValidSyntax = true;
            try
            {
                JsonDocument.Parse(jsonData);
            }
            catch (JsonException)
            {
                isValidSyntax = false;
            }

            // Simple schema validation: check required properties exist and types match
            bool isValidSchema = false;
            if (isValidSyntax)
            {
                try
                {
                    using JsonDocument doc = JsonDocument.Parse(jsonData);
                    JsonElement root = doc.RootElement;

                    bool hasName = root.TryGetProperty("Name", out JsonElement nameElem) && nameElem.ValueKind == JsonValueKind.String;
                    bool hasAge = root.TryGetProperty("Age", out JsonElement ageElem) && ageElem.ValueKind == JsonValueKind.Number;
                    bool hasEmail = root.TryGetProperty("Email", out JsonElement emailElem) && emailElem.ValueKind == JsonValueKind.String;

                    isValidSchema = hasName && hasAge && hasEmail;
                }
                catch
                {
                    isValidSchema = false;
                }
            }

            // Output validation result
            if (isValidSyntax && isValidSchema)
                Console.WriteLine("JSON is valid according to the schema.");
            else if (!isValidSyntax)
                Console.WriteLine("JSON is syntactically invalid.");
            else
                Console.WriteLine("JSON does NOT match the schema.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
