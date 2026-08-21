// Title: Export Workbook ContentTypeProperty Names to CSV with Aspose.Cells for .NET (C#)
// Description: This C# example uses Aspose.Cells to open an existing .xlsx file, reads every ContentTypeProperty defined in the workbook, and writes the property names to a CSV file named ContentTypePropertiesReport.csv. The code includes file‑existence checking, proper CSV quoting, and basic exception handling.
// Keywords: Aspose.Cells | C# | .NET | ContentTypeProperty | export to CSV | Excel custom properties | workbook metadata extraction | CSV report | GitHub example | US developers | European developers
// Common Searches: How to list ContentTypeProperty names from an Excel file using Aspose.Cells C# | Aspose.Cells export workbook custom properties to CSV | C# code to write Excel metadata to CSV | Sample project ExportContentTypePropertyNamesToCsv on GitHub | Extract Excel content type metadata with Aspose.Cells
// Developer Intent: Generate a CSV inventory of all ContentTypeProperty names from a workbook.
// Use Cases: Compliance audit of custom content‑type fields across spreadsheets | Pre‑migration metadata snapshot before moving files to a new system | Feeding property names into downstream validation or ETL pipelines | Creating quick documentation for business analysts
// AI Prompts: Provide C# code that loads an .xlsx file with Aspose.Cells, iterates workbook.ContentTypeProperties, and saves each property name to a CSV file with proper quoting. | Show how to handle missing input files and exceptions when exporting ContentTypeProperty names in a .NET console app. | Explain how to escape double quotes in CSV output for property names using C# and Aspose.Cells. | Suggest ways to integrate this export routine into an automated reporting workflow on Windows or Azure.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    // This C# example uses Aspose.Cells to open an existing .xlsx file, reads every ContentTypeProperty defined in the workbook, and writes the property names to a CSV file named ContentTypePropertiesReport.csv. The code includes file‑existence checking, proper CSV quoting, and basic exception handling.
    public class ExportContentTypePropertyNamesToCsv
    {
        // Entry point for the application
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook from the existing file
                Workbook workbook = new Workbook(inputPath);

                // Prepare the CSV file path
                string csvPath = "ContentTypePropertiesReport.csv";

                // Write property names to CSV
                using (StreamWriter writer = new StreamWriter(csvPath))
                {
                    writer.WriteLine("PropertyName");

                    foreach (ContentTypeProperty property in workbook.ContentTypeProperties)
                    {
                        string name = property.Name ?? string.Empty;
                        string escapedName = name.Replace("\"", "\"\"");
                        writer.WriteLine($"\"{escapedName}\"");
                    }
                }

                Console.WriteLine($"Content type property names have been exported to '{csvPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
