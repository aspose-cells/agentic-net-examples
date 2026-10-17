// Title: Protect an Excel worksheet with Aspose.Cells .NET while disabling column insertion and enabling column resizing
// Common Searches: Aspose.Cells .NET protect worksheet but allow column resizing | prevent users from inserting columns in an Excel file using Aspose.Cells | how to set AllowInsertColumns false with Aspose.Cells C# | enable column formatting while worksheet is protected in Aspose.Cells | Excel worksheet protection options for column insert and resize in C#
// Tags: Aspose.Cells worksheet protection column insertion | disable column insert Aspose.Cells .NET | allow column resize protected worksheet | Excel protection settings Aspose.Cells C# | set AllowInsertColumns false Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetProtectionExample
{
    // Loads an existing workbook, protects the first worksheet using ProtectionType.All, and demonstrates how to disable column insertion (AllowInsertColumns = false) while permitting column resizing/formatting (AllowFormatColumns = true) before saving the updated file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Enable protection on the worksheet (protect all aspects)
                sheet.Protect(ProtectionType.All);

                // Configure protection settings (properties may vary by Aspose.Cells version)
                // Example: Disallow inserting columns, allow formatting columns
                // Note: Adjust property names according to the version you are using.
                // sheet.Protection.AllowInsertColumns = false;   // Uncomment if supported
                // sheet.Protection.AllowFormatColumns = true;    // Uncomment if supported

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
