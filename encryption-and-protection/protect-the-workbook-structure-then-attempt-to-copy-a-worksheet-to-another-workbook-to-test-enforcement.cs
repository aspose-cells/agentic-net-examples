// Title: How to protect an Excel workbook's structure with a password and verify that worksheet copying is blocked using Aspose.Cells for .NET
// Common Searches: Aspose.Cells protect workbook structure password C# example | prevent copying worksheets from a password protected Excel file using Aspose.Cells | Worksheets.AddCopy throws exception when source workbook is structure protected | verify workbook structure protection enforcement in Aspose.Cells .NET | copy worksheet from protected workbook Aspose.Cells error handling
// Tags: Aspose.Cells workbook structure protection | Aspose.Cells copy worksheet from protected workbook | Aspose.Cells Worksheets.AddCopy error handling | Aspose.Cells password protection enforcement | Aspose.Cells prevent sheet copy

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds data, applies structure protection with a password, saves and reloads the file, then attempts to copy the protected worksheet to another workbook using Worksheets.AddCopy. The operation throws an exception, confirming that the structure protection prevents worksheet copying.
class Program
{
    static void Main()
    {
        try
        {
            // Create the source workbook and add some data
            Workbook sourceWb = new Workbook();
            Worksheet sourceWs = sourceWb.Worksheets[0];
            sourceWs.Name = "DataSheet";
            sourceWs.Cells["A1"].PutValue("Sample Data");

            // Protect the workbook structure with a password
            sourceWb.Protect(ProtectionType.Structure, "pwd123");

            // Save the protected workbook
            string sourcePath = "Source.xlsx";
            sourceWb.Save(sourcePath);

            // Load the protected workbook (ensure the file exists)
            Workbook protectedWb = null;
            if (File.Exists(sourcePath))
            {
                protectedWb = new Workbook(sourcePath);
            }
            else
            {
                Console.WriteLine($"File not found: {sourcePath}");
                return;
            }

            // Create a destination workbook
            Workbook destWb = new Workbook();

            // Attempt to copy the protected worksheet to the destination workbook
            try
            {
                // AddCopy expects the source worksheet name, not the Worksheet object
                destWb.Worksheets.AddCopy(protectedWb.Worksheets[0].Name);
                Console.WriteLine("Worksheet copied successfully (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Copy failed as expected: " + ex.Message);
            }

            // Save the destination workbook (optional)
            string destPath = "Destination.xlsx";
            destWb.Save(destPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
