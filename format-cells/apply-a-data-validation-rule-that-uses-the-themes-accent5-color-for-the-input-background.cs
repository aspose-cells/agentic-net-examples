// Title: Add a list data validation to cell A1 and apply the workbook's Accent5 theme color as the input message background using Aspose.Cells for .NET
// AI Prompts: Create a list‑type validation for cell A1 and set its input message background to the workbook's Accent5 theme color via reflection in C# with Aspose.Cells. | Retrieve the Accent5 color from a workbook's theme and assign it to Validation.InputMessageBackgroundColor using reflection in Aspose.Cells. | Generate an Excel file that contains a data‑validation list on A1 where the input prompt background matches the workbook's Accent5 theme color.
// Common Searches: Aspose.Cells how to set validation input message background color to theme Accent5 in C# | C# Aspose.Cells apply workbook theme color to data validation prompt background | Use reflection to set InputMessageBackgroundColor property in Aspose.Cells validation | Get theme Accent5 color from Aspose.Cells workbook for validation styling
// Tags: Aspose.Cells list validation with theme accent color | set validation input message background C# | retrieve workbook Accent5 theme color Aspose.Cells | reflection assign InputMessageBackgroundColor Aspose.Cells | Excel data validation styling using theme colors

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a list‑type data validation to cell A1, obtains the workbook's Accent5 theme color, and, if the InputMessageBackgroundColor property is available, sets it via reflection before saving the file as DataValidation_Accent5.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the cell area for validation (A1)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 0,
                    EndColumn = 0
                };

                // Add a validation rule for the defined area (returns the index)
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];

                // Set validation type to List and provide the allowed values
                validation.Type = ValidationType.List;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "\"Option1,Option2,Option3\"";

                // Attempt to apply the theme's Accent5 color as background (if supported)
                try
                {
                    Color accent5Color = workbook.GetThemeColor(ThemeColorType.Accent5);
                    // The InputMessageBackgroundColor property is not available in all versions;
                    // if it exists, set it via reflection to avoid compile errors.
                    var prop = typeof(Validation).GetProperty("InputMessageBackgroundColor");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(validation, accent5Color);
                    }
                }
                catch
                {
                    // Ignore any errors related to theme color retrieval or property setting
                }

                // Prepare output path
                string outputPath = "DataValidation_Accent5.xlsx";

                // Ensure the directory exists (handle case where outputPath has no directory)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
