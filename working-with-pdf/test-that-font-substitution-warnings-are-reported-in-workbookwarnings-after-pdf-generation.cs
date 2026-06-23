using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Saving;   // PdfSaveOptions resides in this namespace

namespace FontSubstitutionWarningDemo
{
    // Custom warning callback that stores all warnings in a list
    public class CollectingWarningCallback : IWarningCallback
    {
        public List<WarningInfo> CollectedWarnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            // Store every warning for later inspection
            CollectedWarnings.Add(warningInfo);
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put some text into a cell and assign a font that does not exist on the system
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Text with a missing font");
                Style style = workbook.CreateStyle();
                style.Font.Name = "NonExistentFont"; // Expected to trigger substitution
                cell.SetStyle(style);

                // Attach the custom warning callback to capture warnings during PDF generation
                var warningCallback = new CollectingWarningCallback();
                workbook.Settings.WarningCallback = warningCallback;

                // Save the workbook as PDF (triggers rendering and possible font substitution)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                workbook.Save("output.pdf", pdfOptions);

                // After saving, inspect the collected warnings for font substitution entries
                int fontSubstitutionCount = 0;
                foreach (var warning in warningCallback.CollectedWarnings)
                {
                    if (warning.WarningType == WarningType.FontSubstitution)
                    {
                        fontSubstitutionCount++;
                        Console.WriteLine($"Font substitution warning: {warning.Description}");
                    }
                }

                Console.WriteLine($"Total font substitution warnings captured: {fontSubstitutionCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}