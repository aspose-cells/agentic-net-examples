// Title: Load an OTS spreadsheet template, replace placeholder text, and save as ODS using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example for loading an OTS template and exporting to ODS | how to replace {Name} placeholder in an OTS file with Aspose.Cells .NET | convert OTS spreadsheet to ODS after text substitution using Aspose.Cells C# | Aspose.Cells code to modify OTS template and save as ODS format in .NET
// Tags: load OTS workbook Aspose.Cells | placeholder substitution Aspose.Cells C# | save workbook as ODS Aspose.Cells | template driven ODS export .NET | OTS to ODS conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program demonstrates how to open an OTS spreadsheet template with Aspose.Cells, replace a placeholder token such as {Name} with a real value, and then save the updated workbook as an ODS file.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.ots";
            const string resultPath = "result.ods";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the OTS template file
            var workbook = new Workbook(templatePath);

            // Replace placeholder text with actual values using the simple overload
            workbook.Replace("{Name}", "John Doe");

            // Additional replacements can be added similarly, e.g.:
            // workbook.Replace("{Date}", DateTime.Now.ToShortDateString());

            // Save the modified workbook as an ODS file
            workbook.Save(resultPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
