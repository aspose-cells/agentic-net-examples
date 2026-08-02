using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class WorkbookWithPasswordQr
{
    static void Main()
    {
        try
        {
            // Define the workbook protection password
            string password = "SecurePass123";

            // Create a new workbook and set the opening password
            Workbook wb = new Workbook();
            wb.Settings.Password = password;

            // Base64‑encoded 1×1 transparent PNG
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
            byte[] pngData = Convert.FromBase64String(base64Png);

            // Insert the image into the first worksheet at cell A1
            using (MemoryStream ms = new MemoryStream(pngData))
            {
                Worksheet sheet = wb.Worksheets[0];
                int pictureIndex = sheet.Pictures.Add(0, 0, ms);
                Picture pic = sheet.Pictures[pictureIndex];
                pic.Width = 150;   // Desired width
                pic.Height = 150;  // Desired height
            }

            // Ensure the output directory exists
            string outputPath = "Workbook_With_Password_QR.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the protected workbook
            wb.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}