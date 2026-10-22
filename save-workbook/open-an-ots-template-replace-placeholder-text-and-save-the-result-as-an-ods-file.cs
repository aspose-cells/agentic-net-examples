// Title: Load an OTS spreadsheet template, replace placeholders, and save as ODS using Aspose.Cells for .NET
// AI Prompts: Load an OTS file with Aspose.Cells, replace the {Name} placeholder with a custom value, and export the workbook to ODS. | Extend the example to replace multiple placeholders such as {Name} and {Date} in an OTS template before saving it as ODS. | Add robust error handling that verifies the template file exists, logs a friendly warning if it is missing, and catches exceptions during OTS‑to‑ODS conversion.
// Common Searches: Aspose.Cells C# replace placeholder text in OTS template and save as ODS | How to edit an OpenDocument spreadsheet template (OTS) with Aspose.Cells .NET | Convert OTS file to ODS after performing text substitution using Aspose.Cells | C# example for loading OTS workbook, replacing {Name}, and exporting to ODS
// Tags: placeholder replacement in OTS workbook Aspose.Cells | export edited OTS to ODS Aspose.Cells | load OTS template with Aspose.Cells C# | text substitution in OpenDocument spreadsheet Aspose.Cells | file existence check for OTS template Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that a template.ots file exists, loads it into an Aspose.Cells Workbook, replaces the {Name} placeholder with "John Doe", and then saves the modified workbook as result.ods using the ODS save format, while handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.ots";
            const string resultPath = "result.ods";

            // Verify that the template file exists before attempting to load it
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the OTS template workbook
            Workbook workbook = new Workbook(templatePath);

            // Replace placeholders throughout the workbook
            workbook.Replace("{Name}", "John Doe");
            // Additional replacements can be added similarly:
            // workbook.Replace("{Date}", DateTime.Today.ToShortDateString());

            // Save the modified workbook as ODS
            workbook.Save(resultPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
