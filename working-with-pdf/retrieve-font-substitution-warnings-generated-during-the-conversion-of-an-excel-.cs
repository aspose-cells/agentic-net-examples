using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace FontSubstitutionWarningDemo
{
    // Custom warning callback to capture font substitution warnings
    public class FontWarningCallback : IWarningCallback
    {
        // List to store warning messages
        public List<string> FontSubstitutionWarnings { get; } = new List<string>();

        // This method is called by Aspose.Cells during rendering/saving
        public void Warning(WarningInfo warningInfo)
        {
            // Capture only font substitution warnings
            if (warningInfo.WarningType == WarningType.FontSubstitution)
            {
                FontSubstitutionWarnings.Add(warningInfo.Description);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put some text that will require a font not present on the system
            sheet.Cells["A1"].PutValue("Text with a missing font");

            // Apply a non‑existent font to trigger substitution
            Style style = workbook.CreateStyle();
            style.Font.Name = "NonExistentFont";
            sheet.Cells["A1"].SetStyle(style);

            // Prepare PDF save options and attach the custom warning callback
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            FontWarningCallback warningCallback = new FontWarningCallback();
            saveOptions.WarningCallback = warningCallback;

            // Save the workbook to PDF (conversion)
            workbook.Save("output.pdf", saveOptions);

            // Output captured font substitution warnings
            Console.WriteLine("Font Substitution Warnings:");
            if (warningCallback.FontSubstitutionWarnings.Count == 0)
            {
                Console.WriteLine("No font substitution warnings were generated.");
            }
            else
            {
                foreach (string msg in warningCallback.FontSubstitutionWarnings)
                {
                    Console.WriteLine("- " + msg);
                }
            }
        }
    }
}