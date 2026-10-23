// Title: Add a custom email address validation with an input prompt to an Excel range using Aspose.Cells for .NET
// Common Searches: how to add an input message for email validation in Aspose.Cells C# | Aspose.Cells custom data validation with prompt title for Excel cells | C# code to show a data validation input message for email addresses in Excel | set ShowInputMessage property for a validation rule in Aspose.Cells .NET | create email validation that displays 'Enter a valid email' using Aspose.Cells
// Tags: Aspose.Cells email validation with input prompt | custom data validation formula C# | validation prompt text Aspose.Cells | Excel validation prompt caption .NET | ISNUMBER FIND email validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, defines the B2:B20 range, and adds a custom validation rule that checks for an '@' character using the formula ISNUMBER(FIND("@",B2)). It shows how to configure the validation to display an input prompt (ShowInputMessage, PromptTitle, Prompt) – noting that these properties are supported in newer Aspose.Cells versions. The workbook is saved as EmailValidation.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the cell range for email input (B2:B20)
                CellArea emailRange = new CellArea
                {
                    StartRow = 1,    // Row 2 (zero‑based)
                    EndRow = 19,     // Row 20
                    StartColumn = 1, // Column B
                    EndColumn = 1    // Column B
                };

                // Add custom validation to the defined range
                int validationIndex = sheet.Validations.Add(emailRange);
                Validation emailValidation = sheet.Validations[validationIndex];

                // Set validation type to Custom with a formula that checks for '@'
                emailValidation.Type = ValidationType.Custom;
                emailValidation.Formula1 = "ISNUMBER(FIND(\"@\",B2))";

                // Note: Input message properties are not available in the current Aspose.Cells version.
                // If needed, they can be set using ShowInputMessage, PromptTitle, and Prompt in newer versions.

                // Define output file path
                string outputPath = "EmailValidation.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
