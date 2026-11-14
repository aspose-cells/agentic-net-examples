// Title: Apply write protection with author name and audit comments to an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to protect a workbook and store author information | set user name and comments on Excel workbook protection with Aspose.Cells | apply write‑protect to entire workbook and add audit log in .NET | using reflection to set Protection properties in Aspose.Cells C# | save password‑protected Excel file with metadata using Aspose.Cells
// Tags: workbook password protection Aspose.Cells | set protection user name Aspose.Cells | add audit comments to protected Excel workbook C# | protect entire Excel file Aspose.Cells .NET | reflection access Protection object Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, inserts sample data, applies password‑based write protection to the whole workbook, and—via reflection—sets the protection's user name and audit comments when the Protection API is available, then saves the file while ensuring the output directory exists.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B1"].PutValue(100);

            // Define a password for write protection
            string password = "MySecurePassword";

            // Apply write protection to the entire workbook
            workbook.Protect(ProtectionType.All, password);

            // Set audit information if the Protection API is available (newer versions)
            var protectionProp = workbook.GetType().GetProperty("Protection");
            if (protectionProp != null)
            {
                // Retrieve the Protection object via reflection
                object protectionObj = protectionProp.GetValue(workbook);
                if (protectionObj != null)
                {
                    // Use dynamic to invoke methods without compile‑time binding
                    dynamic protection = protectionObj;
                    protection.SetUserName("Jane Smith");
                    protection.SetComments("Workbook write‑protected for audit tracking on " + DateTime.Now.ToString("yyyy-MM-dd"));
                }
            }

            // Save the protected workbook
            string outputPath = "WriteProtectedWorkbook.xlsx";

            // Ensure the directory exists (prevents FileNotFoundException for invalid paths)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
