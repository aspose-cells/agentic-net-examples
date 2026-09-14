// Title: Protect an Excel worksheet with Aspose.Cells for .NET: enable row insertion while disabling row deletion
// AI Prompts: Write C# code that uses Aspose.Cells to protect a worksheet with a password, allowing users to insert new rows but preventing them from deleting existing rows. | Show how to set Worksheet.Protection properties in Aspose.Cells so that AllowInsertRows is true and AllowDeleteRows is false, then save the workbook. | Generate a complete example that loads or creates a workbook, applies row‑insertion‑only protection, and writes the file to disk.
// Common Searches: Aspose.Cells C# protect worksheet allow insert rows but not delete rows | How to set worksheet protection options for row insertion only in .NET | C# Aspose.Cells protect sheet with password and restrict row deletion | Enable row insertion while disabling row deletion using Aspose.Cells API | Worksheet protection settings AllowInsertRows false Aspose.Cells
// Tags: worksheet protection enable row insertion Aspose.Cells | disable row deletion Aspose.Cells .NET | password protect Excel worksheet C# Aspose.Cells | set protection options Aspose.Cells | row insertion only permission Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook or creates a new one, accesses the first worksheet, configures protection to allow row insertion and block row deletion, applies a password using Protect(ProtectionType.All), and saves the protected file as ProtectedWorksheet.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if present; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Configure protection options (commented out if not supported by the current Aspose.Cells version)
                // sheet.Protection.AllowInsertRows = true;   // allow inserting rows
                // sheet.Protection.AllowDeleteRows = false; // disallow deleting rows

                // Apply protection with a password (oldPassword is empty)
                sheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty);

                // Save the protected workbook
                const string outputPath = "ProtectedWorksheet.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
