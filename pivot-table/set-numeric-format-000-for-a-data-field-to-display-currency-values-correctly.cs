// Title: How to format an Excel cell as currency with the '#,##0.00' custom numeric format using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom numeric format for a cell | display currency in Excel workbook using '#,##0.00' with Aspose.Cells | apply number style to a specific cell in .NET Excel file | format cell value as currency programmatically with Aspose.Cells | C# example of setting custom number format in Excel worksheet
// Tags: Aspose.Cells cell number format customization | currency display format '#,##0.00' in Excel via .NET | programmatic cell style application Aspose.Cells | Excel worksheet cell formatting C# Aspose | set custom number pattern for Excel cell using Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, writes a numeric value to cell A1, retrieves its style, assigns the custom numeric format '#,##0.00' to display the value as currency, reapplies the style, and saves the file as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = CreateWorkbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Put a sample numeric value into cell A1
                Cell cell = sheet.Cells["A1"];
                cell.PutValue(1234.56);

                // Retrieve the cell's current style
                Style style = cell.GetStyle();

                // Set a custom numeric format to display currency values
                style.Custom = "#,##0.00";

                // Apply the updated style back to the cell
                cell.SetStyle(style);

                // Save the workbook to the specified file
                SaveWorkbook(workbook, "output.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a new workbook; replace with loading logic if a template file is required
        private static Workbook CreateWorkbook()
        {
            // Example of loading a template if it exists:
            // string templatePath = "template.xlsx";
            // if (File.Exists(templatePath))
            //     return new Workbook(templatePath);
            // else
            //     return new Workbook();

            return new Workbook();
        }

        // Saves the workbook to the given path, ensuring the directory exists
        private static void SaveWorkbook(Workbook workbook, string filePath)
        {
            try
            {
                string directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
                throw;
            }
        }
    }
}
