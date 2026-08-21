// Title: Batch Apply a Holiday Theme to Multiple Excel Workbooks with Aspose.Cells for .NET
// Description: C# program that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, defines a 12‑color holiday palette, applies it as a custom theme named "HolidayTheme", and saves the themed files to a separate output directory while handling missing files and runtime errors.
// Keywords: Aspose.Cells custom theme | C# batch Excel theme | holiday color palette Excel | apply theme to multiple workbooks | bulk Excel formatting .NET | CustomTheme method Aspose | automate Excel styling | seasonal Excel report template
// Common Searches: how to apply a custom theme to all Excel files in a folder using Aspose.Cells | batch process to set holiday colors in multiple workbooks .NET | apply 12‑color palette to several Excel workbooks programmatically | save themed Excel files to a different directory with Aspose.Cells | bulk update Excel workbook themes C#
// Developer Intent: Programmatically add a predefined holiday color scheme to every workbook in a specified directory and write the themed copies to an output folder.
// Use Cases: Prepare a festive report package where each workbook shares the same holiday theme before distribution. | Standardize corporate branding across dozens of Excel templates by applying a uniform custom theme in bulk. | Automate the creation of year‑end client deliverables with a seasonal color scheme to enhance visual appeal.
// AI Prompts: Write C# code that reads all .xlsx files from a folder and applies a custom holiday theme using Aspose.Cells, including robust error handling. | Show how to define a 12‑color CustomTheme for a holiday palette and apply it to each workbook in a batch process. | Explain how to modify the batch routine to target only specific worksheets within each workbook when applying the custom theme.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// C# program that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, defines a 12‑color holiday palette, applies it as a custom theme named "HolidayTheme", and saves the themed files to a separate output directory while handling missing files and runtime errors.
class HolidayThemeBatch
{
    static void Main()
    {
        // Folder containing the source workbooks
        string inputFolder = @"C:\Workbooks\Input";
        // Folder where the themed workbooks will be saved
        string outputFolder = @"C:\Workbooks\Output";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Define a holiday theme (12 colors as required by CustomTheme)
        Color[] holidayColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1 - white
            Color.FromArgb(0, 0, 0),       // Text1 - black
            Color.FromArgb(255, 255, 255), // Background2 - white
            Color.FromArgb(0, 0, 0),       // Text2 - black
            Color.FromArgb(255, 0, 0),     // Accent1 - red
            Color.FromArgb(0, 128, 0),     // Accent2 - dark green
            Color.FromArgb(255, 215, 0),   // Accent3 - gold
            Color.FromArgb(255, 165, 0),   // Accent4 - orange
            Color.FromArgb(0, 0, 255),     // Accent5 - blue
            Color.FromArgb(128, 0, 128),   // Accent6 - purple
            Color.FromArgb(0, 0, 255),     // Hyperlink - blue
            Color.FromArgb(255, 0, 0)      // Followed Hyperlink - red
        };

        // Iterate over each Excel file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Apply the custom holiday theme
                workbook.CustomTheme("HolidayTheme", holidayColors);

                // Build the output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the themed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Holiday theme applied to all workbooks successfully.");
    }
}
