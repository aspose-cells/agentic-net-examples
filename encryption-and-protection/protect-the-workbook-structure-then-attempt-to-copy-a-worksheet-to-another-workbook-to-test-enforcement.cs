// Title: Protect workbook structure and verify that worksheet copying is blocked with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells protect workbook structure C# prevent AddCopy | C# code to lock workbook structure and catch copy restriction exception using Aspose.Cells | How to test structure protection by trying to copy a worksheet in Aspose.Cells for .NET | Worksheet.AddCopy fails when source workbook is structure protected Aspose.Cells | Example of protecting workbook without password and handling copy error in C#
// Tags: Aspose.Cells workbook structure protection C# | AddCopy restriction on protected workbook | exception handling for structure locked workbook Aspose.Cells | protect workbook without password Aspose.Cells | verify copy block using Worksheets.AddCopy

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStructureProtectionDemo
{
    // The sample creates a source workbook, applies structure protection with an empty password, and then attempts to copy the protected worksheet to a new workbook using Worksheets.AddCopy. The expected exception is caught, demonstrating that structure protection prevents worksheet copying.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a source workbook and add a worksheet with some data
                Workbook sourceWorkbook = new Workbook();
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                sourceSheet.Name = "SourceSheet";
                sourceSheet.Cells["A1"].PutValue("Hello, Aspose!");

                // Protect the workbook structure (empty password)
                sourceWorkbook.Protect(ProtectionType.Structure, string.Empty);

                // Save the protected workbook (optional, just for demonstration)
                string protectedPath = "ProtectedSource.xlsx";
                sourceWorkbook.Save(protectedPath);

                // Create a destination workbook where we will attempt to copy the protected sheet
                Workbook destWorkbook = new Workbook();

                try
                {
                    // Attempt to copy the protected worksheet to the destination workbook
                    // This operation should fail because the source workbook's structure is protected
                    destWorkbook.Worksheets.AddCopy(sourceSheet.Name);
                    Console.WriteLine("Worksheet copied successfully (unexpected).");
                }
                catch (Exception ex)
                {
                    // Expected path: an exception is thrown due to structure protection
                    Console.WriteLine("Failed to copy worksheet as expected:");
                    Console.WriteLine(ex.Message);
                }

                // Save the destination workbook (will contain no copied sheet if protection worked)
                string destPath = "Destination.xlsx";
                destWorkbook.Save(destPath);
            }
            catch (Exception e)
            {
                // General exception handling for unexpected errors
                Console.WriteLine("An error occurred:");
                Console.WriteLine(e.Message);
            }
        }
    }
}
