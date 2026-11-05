// Title: Password‑protect an Excel workbook’s structure and windows and save it as .xlsx using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect workbook structure and windows with password and save as xlsx | how to create a password‑protected Excel file using Aspose.Cells .NET | Aspose.Cells Workbook.Protect example for disabling changes before export | save a protected workbook to the default .xlsx format with Aspose.Cells in C# | protect all workbook elements and write to .xlsx using Aspose.Cells for .NET
// Tags: protect workbook structure Aspose.Cells C# | password protect Excel workbook Aspose.Cells | save protected workbook as xlsx Aspose.Cells | Workbook.Protect API usage Aspose.Cells | disable workbook changes before saving C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to apply full protection (structure and windows) with a password to an Excel workbook using Aspose.Cells for .NET and then save the protected file in the default .xlsx format.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Protect the workbook (structure and windows) with a password
                workbook.Protect(ProtectionType.All, "StrongPassword123");

                // Save the protected workbook
                string outputPath = "ProtectedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
