// Title: Saving an Aspose.Cells workbook as ODS with ODF 1.2 compliance in C#
// Common Searches: C# Aspose.Cells example to export Excel workbook to ODS with ODF 1.2 | How to use OdsSaveOptions to create ODS file in .NET | Aspose.Cells ODF 1.2 ODS conversion code sample C# | Generate ODS spreadsheet from Aspose.Cells with OdfVersion.Version12 | Save workbook as ODS format using Aspose.Cells library in C#
// Tags: Aspose.Cells OdsSaveOptions ODF 1.2 | C# export workbook to ODS | ODF 1.2 ODS conversion Aspose.Cells | .NET generate ODS file | Aspose.Cells ODS save options example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, fills it with data, configures OdsSaveOptions (defaulting to ODF 1.2), ensures the output directory exists, and saves the workbook as an ODS file using Aspose.Cells in C#, with error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Configure ODS save options (default is ODF 1.2)
            OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Ods);

            // Define output path and ensure the directory exists
            string outputPath = "SampleOutput.ods";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an ODS file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
