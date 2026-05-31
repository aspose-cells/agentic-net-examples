using System;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using Aspose.Cells;

class HtmlToExcelWithBackground
{
    static void Main()
    {
        try
        {
            // Paths for input HTML and output Excel files
            string htmlFilePath = "input.html";
            string excelFilePath = "output.xlsx";

            // Verify that the HTML source file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlFilePath}'.");
                return;
            }

            // Load the HTML content as a string
            string htmlContent = File.ReadAllText(htmlFilePath);

            // Regex to capture a URL from a CSS background-image declaration
            string pattern = @"background-image\s*:\s*url\(['""]?(?<url>[^'"")]+)['""]?\)";
            Match match = Regex.Match(htmlContent, pattern, RegexOptions.IgnoreCase);

            // Download the background image (if any) to a temporary file
            string tempImagePath = null;
            if (match.Success)
            {
                string imageUrl = match.Groups["url"].Value;
                string extension = Path.GetExtension(imageUrl);
                tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + extension);

                using (HttpClient client = new HttpClient())
                {
                    byte[] data = client.GetByteArrayAsync(imageUrl).Result;
                    File.WriteAllBytes(tempImagePath, data);
                }
            }

            // Load the HTML file into a workbook (Aspose.Cells can load HTML directly)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // If a background image was downloaded, add it as a picture covering the sheet
            if (tempImagePath != null && File.Exists(tempImagePath))
            {
                Worksheet sheet = workbook.Worksheets[0];

                // Add picture at the top‑left corner; it will act as a background visual
                sheet.Pictures.Add(0, 0, tempImagePath);

                // Delete the temporary image file
                File.Delete(tempImagePath);
            }

            // Save the workbook as an Excel file
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Excel file saved to: {excelFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}