// Title: Export an Aspose.Cells workbook to a semicolon‑delimited CSV file using TxtSaveOptions in C#
// AI Prompts: Write C# code that creates a workbook with sample data and saves it as a CSV file using a semicolon as the column separator via Aspose.Cells TxtSaveOptions. | Show how to set TxtSaveOptions.Separator to ';' and ensure the target directory exists before calling Workbook.Save in Aspose.Cells for .NET. | Demonstrate error handling while exporting a workbook to a semicolon‑delimited CSV with Aspose.Cells, including folder creation and console output.
// Common Searches: Aspose.Cells C# export workbook to CSV with custom delimiter semicolon | How to set TxtSaveOptions separator property for CSV in Aspose.Cells .NET | Save workbook as semicolon separated values using Aspose.Cells | Create output folder automatically when saving CSV with Aspose.Cells
// Tags: Aspose.Cells TxtSaveOptions CSV delimiter | semicolon delimited CSV export .NET | Workbook.Save custom separator Aspose | ensure output directory exists C# Aspose.Cells | export workbook to CSV Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook, adding data, configuring TxtSaveOptions.Separator to ';', ensuring the output folder exists, and saving the workbook as a semicolon‑delimited CSV file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);

            // Set CSV save options to use semicolon as column delimiter
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
            csvOptions.Separator = ';';

            // Define output file path
            string outputPath = "output.csv";

            // Ensure the directory exists (in case a relative path is used)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as CSV with the custom delimiter
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
