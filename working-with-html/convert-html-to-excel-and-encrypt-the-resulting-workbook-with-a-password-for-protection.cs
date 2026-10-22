// Title: Convert an HTML string containing a table to a password‑protected XLSX file using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an HTML string, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, sets a password, and saves the workbook as an encrypted XLSX file. | Show how to protect an Excel workbook generated from HTML by assigning Workbook.Settings.Password before calling Workbook.Save with Aspose.Cells.
// Common Searches: c# asp.net convert html table to encrypted xlsx using aspose.cells | how to set workbook password when saving html to excel with aspose.cells | load html from string into workbook and protect with password asp.net | aspose.cells encrypt excel file created from html content
// Tags: html to xlsx conversion Aspose.Cells | apply password protection Aspose.Cells workbook | load html stream Aspose.Cells C# | save encrypted xlsx Aspose.Cells | protect workbook generated from html Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an HTML string containing a table into a MemoryStream, uses Aspose.Cells HtmlLoadOptions to create a Workbook, assigns a password via Workbook.Settings.Password, and saves the workbook as an encrypted XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Sample HTML string to be converted.
            string htmlContent = @"
                <html>
                    <body>
                        <table border='1'>
                            <tr><th>Name</th><th>Age</th></tr>
                            <tr><td>Alice</td><td>30</td></tr>
                            <tr><td>Bob</td><td>25</td></tr>
                        </table>
                    </body>
                </html>";

            // Convert HTML string to a memory stream.
            byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Load the HTML content into a new workbook using HtmlLoadOptions.
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Set a password to encrypt the workbook.
                workbook.Settings.Password = "MySecretPassword";

                // Save the encrypted workbook to an XLSX file.
                string outputPath = "EncryptedFromHtml.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
