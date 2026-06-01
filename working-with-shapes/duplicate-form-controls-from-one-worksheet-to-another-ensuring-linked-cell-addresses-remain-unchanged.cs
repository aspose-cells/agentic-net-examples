using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormControlCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "SourceWorkbook.xlsx";
                const string destPath = "DestinationWorkbook.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook that contains the form controls
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0]; // assume first sheet has the controls

                // Create a new workbook that will receive the copied sheet
                Workbook destWorkbook = new Workbook();

                // Add a new empty worksheet to the destination workbook and name it
                Worksheet destSheet = destWorkbook.Worksheets.Add("CopiedSheet");

                // Configure copy options:
                // ReferToDestinationSheet = false ensures that any linked cell references
                // (including those used by form controls) keep pointing to the original cells.
                CopyOptions copyOptions = new CopyOptions
                {
                    ReferToDestinationSheet = false
                };

                // Copy the source worksheet (including form controls) to the destination worksheet
                destSheet.Copy(sourceSheet, copyOptions);

                // Save the result
                destWorkbook.Save(destPath);
                Console.WriteLine($"Workbook copied successfully to {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}