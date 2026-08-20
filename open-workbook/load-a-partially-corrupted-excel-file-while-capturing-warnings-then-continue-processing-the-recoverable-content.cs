// Title: Load a Corrupted Excel Workbook in C# with Aspose.Cells – Capture Warnings and Recover Data
// Description: This example shows how to open a partially damaged .xlsx file using Aspose.Cells LoadOptions, register a custom IWarningCallback to collect load warnings, ignore unnecessary shapes, bypass data‑validation checks, handle the FileCorrupted CellsException to obtain any recoverable Workbook, enable RepairLoad for further operations, iterate through worksheets to read available cells, and save the repaired file while listing all captured warnings.
// Keywords: Aspose.Cells corrupted workbook | C# load damaged Excel | capture load warnings | custom IWarningCallback | ignore useless shapes | disable data validation | RepairLoad option | recover Excel data | CellsException FileCorrupted | save recovered workbook
// Common Searches: how to open a corrupted xlsx with Aspose.Cells | Aspose.Cells load options for damaged Excel files | collect warnings while loading Excel in .NET | recover data from partially corrupted workbook Aspose | ignore shapes when loading corrupted Excel C#
// Developer Intent: Open a partially corrupted Excel file, log all load warnings, and continue processing the recoverable content.
// Use Cases: Log every warning generated during workbook loading for audit or debugging. | Extract a Workbook object from a FileCorrupted exception and read whatever data is still intact. | Enable RepairLoad to allow further manipulation and then save the repaired workbook to a new file.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions with a custom IWarningCallback to open a corrupted .xlsx, capture warnings, and continue with the recovered Workbook. | Explain how to retrieve the Workbook from a CellsException when the file is corrupted and how to activate RepairLoad for subsequent operations. | Show the LoadOptions settings needed to ignore useless shapes and skip data‑validation checks while loading a damaged Excel workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCorruptLoadDemo
{
    // Custom warning callback that stores all warnings for later inspection
    // This example shows how to open a partially damaged .xlsx file using Aspose.Cells LoadOptions, register a custom IWarningCallback to collect load warnings, ignore unnecessary shapes, bypass data‑validation checks, handle the FileCorrupted CellsException to obtain any recoverable Workbook, enable RepairLoad for further operations, iterate through worksheets to read available cells, and save the repaired file while listing all captured warnings.
    public class CustomWarningCallback : IWarningCallback
    {
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            // Store the warning
            Warnings.Add(warningInfo);
            // Optionally, write to console for immediate feedback
            Console.WriteLine($"Warning: {warningInfo.Description}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the partially corrupted Excel file
            string inputPath = "corrupted.xlsx";

            // Initialize the custom warning callback
            var warningCallback = new CustomWarningCallback();

            // Configure load options to capture warnings
            LoadOptions loadOptions = new LoadOptions
            {
                WarningCallback = warningCallback,
                // Ignoring useless shapes can help with corrupted files that contain junk shapes
                IgnoreUselessShapes = true,
                // Continue loading even if data validation errors are present
                CheckDataValid = false
            };

            Workbook workbook = null;

            try
            {
                // Load the workbook with the specified options
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
            {
                // The file is corrupted but Aspose.Cells may have loaded recoverable parts
                Console.WriteLine("FileCorrupted exception caught. Attempting to continue with recovered content.");
                // The workbook object may still be partially loaded; proceed if not null
                if (ex.Data.Contains("Workbook") && ex.Data["Workbook"] is Workbook recoveredWorkbook)
                {
                    workbook = recoveredWorkbook;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
                return;
            }

            if (workbook == null)
            {
                Console.WriteLine("Workbook could not be loaded.");
                return;
            }

            // Enable repair mode for subsequent operations
            workbook.Settings.RepairLoad = true;

            // Example processing: iterate through worksheets and print first cell values
            Console.WriteLine("\nRecoverable content:");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");
                // Attempt to read the value of cell A1; if it fails, catch and continue
                try
                {
                    var cellValue = sheet.Cells["A1"].Value;
                    Console.WriteLine($"  A1 = {cellValue}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Unable to read A1: {ex.Message}");
                }
            }

            // Optionally, save the recovered workbook to a new file
            string outputPath = "recovered_output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nRecovered workbook saved to: {outputPath}");

            // Display all captured warnings
            Console.WriteLine("\nCaptured warnings:");
            foreach (var warning in warningCallback.Warnings)
            {
                Console.WriteLine($"- Type: {warning.Type}, Description: {warning.Description}");
            }
        }
    }
}
