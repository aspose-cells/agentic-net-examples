// Title: Check if an Excel worksheet permits object editing and log the result with Aspose.Cells for .NET (C#)
// Common Searches: how to determine if a worksheet is protected and AllowEditObject is enabled using Aspose.Cells in C# | read AllowEditObject flag from worksheet protection with reflection in Aspose.Cells .NET | log Excel worksheet protection settings for compliance using Aspose.Cells C# example | detect edit objects permission on a protected sheet with Aspose.Cells for .NET | retrieve worksheet protection properties programmatically with Aspose.Cells C#
// Tags: Aspose.Cells worksheet AllowEditObject property | reflection access worksheet protection in .NET | log Excel sheet protection status with Aspose | check object editing permission on protected worksheet | retrieve protection settings from Excel workbook C#

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example loads an Excel workbook using Aspose.Cells, accesses the first worksheet, checks whether the sheet is protected, and uses reflection to read the AllowEditObject flag from the worksheet's Protection object when available. It then writes the protection status and the flag value to the console for compliance logging and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine if the worksheet is protected
            bool isProtected = sheet.IsProtected;

            // Try to read the AllowEditObject setting via reflection (property may not exist in older versions)
            bool allowEditObjects = false;
            PropertyInfo propInfo = sheet.Protection?.GetType().GetProperty("AllowEditObject", BindingFlags.Public | BindingFlags.Instance);
            if (propInfo != null && propInfo.PropertyType == typeof(bool))
            {
                allowEditObjects = (bool)propInfo.GetValue(sheet.Protection);
            }

            // Log the compliance result
            if (isProtected)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' is protected. Allow edit objects: {allowEditObjects}");
            }
            else
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' is not protected.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional, no modifications made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
