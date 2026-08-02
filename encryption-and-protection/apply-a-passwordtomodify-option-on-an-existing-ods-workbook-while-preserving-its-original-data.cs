// Title: C# – Add Write‑Protection Password to an Existing ODS Workbook with Aspose.Cells
// Description: Load an existing ODS file using Aspose.Cells for .NET, assign a modify password via Workbook.Settings.WriteProtection, and save the workbook unchanged so the original data stays intact.
// Keywords: Aspose.Cells C# ODS write protection | set modify password ODS .NET | protect ODS workbook Aspose | Workbook.Settings.WriteProtection example | C# ODS encryption | GitHub Aspose.Cells ODS protection | global spreadsheet security | EU GDPR ODS password
// Common Searches: How to set a modify password on an ODS file with Aspose.Cells C# | Aspose.Cells protect ODS workbook from editing .NET | C# code to add write protection to an existing ODS spreadsheet | Apply password to ODS using Aspose.Cells example | Write‑protect ODS in US compliance
// Developer Intent: Apply a password that blocks editing of an existing ODS workbook while keeping its content unchanged.
// Use Cases: Secure a financial report in ODS format before distributing it to external auditors. | Prevent accidental edits to a shared ODS template used by a multinational team. | Enforce read‑only distribution of a product specification sheet while allowing authorized users to modify it with a password.
// AI Prompts: Generate C# code that loads an existing ODS workbook with Aspose.Cells, sets a write‑protection password, and saves the file preserving all data. | Explain how to programmatically verify that write protection has been applied to an ODS file using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace ApplyWriteProtectionToOdsDemo
{
    // Load an existing ODS file using Aspose.Cells for .NET, assign a modify password via Workbook.Settings.WriteProtection, and save the workbook unchanged so the original data stays intact.
    class ApplyWriteProtectionToOds
    {
        static void Main()
        {
            // Path to the existing ODS workbook
            string inputPath = "ExistingWorkbook.ods";
            string outputPath = "ProtectedWorkbook.ods";

            try
            {
                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    // Load the existing workbook
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    // Create a new workbook as a fallback
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Sheet1";
                    sheet.Cells["A1"].PutValue("Sample data");
                }

                // Apply write‑protection password
                workbook.Settings.WriteProtection.Password = "ModifyPassword123";

                // Save the workbook with write‑protection applied
                workbook.Save(outputPath, SaveFormat.Ods);

                Console.WriteLine($"Write protection applied and workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
