// Title: Apply Aspose.Cells .NET License at Runtime to Remove Evaluation Watermarks (C#)
// Description: This C# example demonstrates how to load an Aspose.Cells.NET.lic file at runtime, activate licensed mode, verify the workbook's licensing status, add sample data, and save the workbook without the evaluation watermark. It also shows graceful handling when the license file is missing.
// Keywords: Aspose.Cells license C# | remove evaluation watermark | Workbook.IsLicensed | runtime license loading | Aspose.Cells .NET example | licensed mode workbook | C# Aspose.Cells tutorial | global Aspose.Cells licensing
// Common Searches: how to set Aspose.Cells license in C# | remove Aspose.Cells evaluation watermark programmatically | check if Aspose.Cells workbook is licensed | load Aspose.Cells .NET license from file | C# example for Aspose.Cells licensed mode
// Developer Intent: Load a valid Aspose.Cells .NET license at runtime to suppress evaluation watermarks and generate a fully licensed workbook.
// Use Cases: Deploy applications that need to run without Aspose.Cells watermarks by applying the license programmatically. | Validate licensing status before using premium features such as advanced charting or PDF conversion. | Provide fallback behavior when the license file is absent, ensuring the app continues to function.
// AI Prompts: Generate C# code that loads an Aspose.Cells license from a given path, checks Workbook.IsLicensed, and saves a workbook without watermarks. | Show error‑handling patterns for missing Aspose.Cells license files while still creating and saving a workbook. | Explain how to verify licensing status after applying a license and enable licensed‑only features conditionally.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to load an Aspose.Cells.NET.lic file at runtime, activate licensed mode, verify the workbook's licensing status, add sample data, and save the workbook without the evaluation watermark. It also shows graceful handling when the license file is missing.
    public class LicensedWorkbookDemo
    {
        public static void Run()
        {
            try
            {
                // Apply the Aspose.Cells license to suppress evaluation watermarks
                string licensePath = "Aspose.Cells.NET.lic";
                if (File.Exists(licensePath))
                {
                    License license = new License();
                    license.SetLicense(licensePath);
                }
                else
                {
                    Console.WriteLine($"License file not found at '{licensePath}'. Continuing without license.");
                }

                // Create a new workbook and verify that the license is active
                Workbook workbook = new Workbook();
                Console.WriteLine($"IsLicensed: {workbook.IsLicensed}");

                // Add some sample content
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";
                sheet.Cells["A1"].PutValue("Licensed Workbook");

                // Save the workbook
                string outputPath = "LicensedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LicensedWorkbookDemo.Run();
        }
    }
}
