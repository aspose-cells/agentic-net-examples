using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define corporate typeface name
                string corporateFont = "CorporateSans";

                // ----- Heading (Major scheme) -----
                Cell headingCell = sheet.Cells["A1"];
                headingCell.PutValue("Corporate Heading");

                // Get the style of the heading cell
                Style headingStyle = headingCell.GetStyle();

                // Set the corporate font name (major scheme concept is implicit)
                headingStyle.Font.Name = corporateFont;   // use Font.Name instead of SetName
                headingStyle.Font.Size = 16;
                headingStyle.Font.IsBold = true;

                // Apply the modified style back to the cell
                headingCell.SetStyle(headingStyle);

                // ----- Body text (Minor scheme) -----
                Cell bodyCell = sheet.Cells["A2"];
                bodyCell.PutValue("Corporate body text goes here.");

                // Get the style of the body cell
                Style bodyStyle = bodyCell.GetStyle();

                // Set the corporate font name (minor scheme concept is implicit)
                bodyStyle.Font.Name = corporateFont;      // use Font.Name instead of SetName
                bodyStyle.Font.Size = 12;

                // Apply the modified style back to the cell
                bodyCell.SetStyle(bodyStyle);

                // Define output file path
                string outputPath = "CorporateThemeFontScheme.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}