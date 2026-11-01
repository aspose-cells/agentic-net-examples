// Title: C# example: generate a JSON configuration file and export an Aspose.Cells workbook to PDF
// Common Searches: how to generate a JSON configuration file in C# before exporting an Aspose.Cells workbook to PDF | Aspose.Cells C# save workbook as PDF with sample data | C# code to create a workbook, add data, and convert it to PDF using Aspose.Cells | using Aspose.Cells SaveFormat.Pdf in a .NET application | example of handling exceptions when saving a workbook to PDF with Aspose.Cells in C#
// Tags: Aspose.Cells save workbook as PDF C# | C# create JSON configuration file | Aspose.Cells workbook to PDF conversion | .NET export Excel to PDF with Aspose.Cells | sample data insertion Aspose.Cells worksheet C# | exception handling Aspose.Cells PDF save

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new Aspose.Cells Workbook, adds sample data to the first worksheet, ensures a simple JSON configuration file exists (creating it if missing), and then saves the workbook as a PDF file named Output.pdf using the SaveFormat.Pdf option. It includes try‑catch blocks to handle errors during file creation and PDF saving.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Path to the JSON configuration file (optional, not attached to PDF in this example)
            string jsonFilePath = "config.json";

            // Ensure the JSON file exists; create a simple one if missing
            if (!File.Exists(jsonFilePath))
            {
                File.WriteAllText(jsonFilePath, "{ \"setting\": \"value\" }");
            }

            // Save the workbook as a PDF
            string outputPdfPath = "Output.pdf";

            try
            {
                workbook.Save(outputPdfPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF successfully saved to '{outputPdfPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
