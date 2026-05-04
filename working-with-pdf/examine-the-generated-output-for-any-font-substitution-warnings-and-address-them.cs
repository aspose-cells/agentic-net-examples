using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace FontSubstitutionDemo
{
    // Custom warning callback to capture font substitution warnings
    public class FontSubstitutionWarningCallback : IWarningCallback
    {
        // List to store warning messages
        public List<string> FontSubstitutionWarnings { get; } = new List<string>();

        // This method is called by Aspose.Cells when a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            // Check if the warning is related to font substitution
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
            // -------------------------------------------------
            // 1. Create a new workbook and access the first sheet
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 2. Add text that uses a non‑existent font to trigger substitution
            // -------------------------------------------------
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Sample text with missing font");
            Style style = workbook.CreateStyle();
            style.Font.Name = "NonExistentFont"; // Font that likely does not exist on the system
            cell.SetStyle(style);

            // -------------------------------------------------
            // 3. Configure font substitution:
            //    - Set system font substitutes preference (optional)
            //    - Define explicit substitutes for the missing font
            // -------------------------------------------------
            FontConfigs.PreferSystemFontSubstitutes = true; // Use OS substitutes first
            FontConfigs.SetFontSubstitutes("NonExistentFont", new string[] { "Arial", "Liberation Sans" });

            // -------------------------------------------------
            // 4. Prepare rendering options with the custom warning callback
            // -------------------------------------------------
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                // Ensure the renderer checks the workbook's default font as a fallback
                CheckWorkbookDefaultFont = true,
                // Attach the warning callback
                WarningCallback = new FontSubstitutionWarningCallback()
            };

            // -------------------------------------------------
            // 5. Render the worksheet to an image (this triggers font processing)
            // -------------------------------------------------
            SheetRender renderer = new SheetRender(sheet, renderOptions);
            renderer.ToImage(0, "RenderedOutput.png");

            // -------------------------------------------------
            // 6. Retrieve and display any captured font substitution warnings
            // -------------------------------------------------
            var callback = (FontSubstitutionWarningCallback)renderOptions.WarningCallback;
            if (callback.FontSubstitutionWarnings.Count > 0)
            {
                Console.WriteLine("Font substitution warnings detected:");
                foreach (string warning in callback.FontSubstitutionWarnings)
                {
                    Console.WriteLine("- " + warning);
                }
            }
            else
            {
                Console.WriteLine("No font substitution warnings were generated.");
            }

            // -------------------------------------------------
            // 7. Save the workbook (optional, to verify that the font settings persist)
            // -------------------------------------------------
            workbook.Save("FontSubstitutionDemo.xlsx");
        }
    }
}