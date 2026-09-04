// Title: Decrypt a password‑protected XLSX workbook and save it as HTML using Aspose.Cells for .NET
// AI Prompts: Create a C# routine that uses Aspose.Cells LoadOptions to open an encrypted .xlsx with a supplied password and then saves the workbook as an HTML file. | Show how to catch a CellsException when the provided password is wrong while loading a protected workbook with Aspose.Cells. | Demonstrate configuring HtmlSaveOptions and exporting a successfully decrypted workbook to HTML, including file‑I/O error handling.
// Common Searches: asp.net load encrypted excel workbook using Aspose.Cells and generate html output | c# Aspose.Cells decrypt protected xlsx and export to html | asp.net handle incorrect password error when opening protected Excel with Aspose.Cells | save workbook as html after verifying password using Aspose.Cells
// Tags: load encrypted xlsx with password Aspose.Cells | export workbook to html Aspose.Cells | handle CellsException incorrect password | HtmlSaveOptions configuration Aspose.Cells | verify workbook password before saving

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies that an encrypted XLSX file exists, loads it with the correct password via LoadOptions, catches password‑related errors, and then converts the decrypted workbook to an HTML document using HtmlSaveOptions, with comprehensive exception handling for both loading and saving steps.
    class Program
    {
        static void Main()
        {
            // Paths and password
            string inputPath = "encrypted.xlsx";
            string password = "myPassword";
            string outputPath = "output.html";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load options with password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            Workbook workbook = null;
            try
            {
                // Attempt to load the encrypted workbook
                workbook = new Workbook(inputPath, loadOptions);
                Console.WriteLine("Password verified and workbook loaded successfully.");
            }
            catch (CellsException ex)
            {
                // Handle incorrect password or other loading issues
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                return;
            }

            try
            {
                // Convert the workbook to HTML
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook converted to HTML at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving HTML: {ex.Message}");
            }
        }
    }
}
