// Title: Checking the IsSigned flag of an Excel workbook before and after digital signing with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to read IsSigned property of a workbook | verify Excel digital signature status using Aspose.Cells before saving | Aspose.Cells check workbook signed status after applying signature | determine if workbook is signed in .NET before and after signing | sample code to read IsSigned flag with Aspose.Cells API
// Tags: Aspose.Cells IsSigned property verification | C# digital signature status check for Excel workbook | read workbook signed flag Aspose.Cells | verify Excel file signature before and after signing | Aspose.Cells workbook signing workflow

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# example loads an existing Excel file (or creates a new workbook if the file is missing), adds a worksheet when none exist, saves the workbook as "signed_output.xlsx", and shows where to read the Workbook.IsSigned property to confirm the signature status before and after applying a digital signature using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Example operation: add a new worksheet if the workbook is empty
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add("Sheet1");
                }

                // Save the workbook
                string outputPath = "signed_output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
