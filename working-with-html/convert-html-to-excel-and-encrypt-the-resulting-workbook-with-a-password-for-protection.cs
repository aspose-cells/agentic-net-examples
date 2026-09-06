// Title: Convert HTML to an Excel workbook and protect it with a password using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook, assigns an opening password, and saves the result as an XLSX file. | Demonstrate how to configure LoadOptions for HTML format and then apply Workbook.Settings.Password to encrypt the Excel output with Aspose.Cells. | Provide a robust example that checks for the existence of the source HTML file and catches exceptions while creating a password‑protected workbook.
// Common Searches: asp.net convert html to excel and set opening password with Aspose.Cells | c# Aspose.Cells load html file and protect workbook with password | how to encrypt generated xlsx from html using Aspose.Cells .NET | sample code for password protecting Excel file after HTML conversion in C#
// Tags: HTML to XLSX conversion Aspose.Cells | Workbook password protection C# | LoadOptions for HTML Aspose.Cells | SaveFormat Xlsx encryption Aspose.Cells | exception handling file existence Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program verifies the presence of an input HTML file, loads it into an Aspose.Cells Workbook using HTML LoadOptions, sets an opening password via Workbook.Settings.Password, and saves the workbook as a password‑protected XLSX file, with exception handling for any errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.xlsx";
            const string password = "MySecurePassword";

            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the HTML file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Encrypt the workbook with a password (required to open the file)
            workbook.Settings.Password = password;

            // Save the workbook as an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
