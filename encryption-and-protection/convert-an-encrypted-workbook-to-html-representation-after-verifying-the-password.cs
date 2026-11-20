// Title: Decrypt a password‑protected Excel workbook and export it to HTML using Aspose.Cells for .NET (C#)
// Common Searches: C# code to open a password‑protected .xlsx with Aspose.Cells and save as HTML | How to load an encrypted Excel file using LoadOptions and convert it to HTML in .NET | Aspose.Cells example for verifying workbook password before exporting to HTML | Convert encrypted Excel workbook to HTML with base64‑encoded images using Aspose.Cells | Handling invalid password exception when converting Excel to HTML in C#
// Tags: load encrypted workbook password Aspose.Cells | export workbook to HTML base64 images | handle CellsException invalid password | verify file existence before loading workbook | HtmlSaveOptions export all worksheets

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample checks that the encrypted Excel file exists, opens it with LoadOptions supplying the password, configures HtmlSaveOptions to embed images as Base64 and export all worksheets, saves the workbook as an HTML file, and catches CellsException to report an invalid password while handling other exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the encrypted workbook and the output HTML file
            string inputFile = @"C:\Path\To\EncryptedWorkbook.xlsx";
            string outputFile = @"C:\Path\To\Workbook.html";

            // Password for the encrypted workbook
            string password = "yourPassword";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // LoadOptions with the password to open the encrypted workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };

                // Load the workbook using the specified LoadOptions
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Set HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = false,
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file using the HtmlSaveOptions overload
                workbook.Save(outputFile, htmlOptions);

                Console.WriteLine("Workbook successfully converted to HTML.");
            }
            catch (CellsException ex)
            {
                // Check for invalid password scenario
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The password provided is invalid.");
                }
                else
                {
                    Console.WriteLine($"CellsException: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Handle other possible exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
