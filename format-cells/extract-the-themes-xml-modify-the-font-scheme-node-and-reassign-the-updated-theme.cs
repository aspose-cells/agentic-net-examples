// Title: Extract Theme XML, Edit Font Scheme, and Reapply Theme with Aspose.Cells for .NET
// Description: Learn how to retrieve the theme XML from an Excel workbook, modify the <a:fontScheme> node to set custom major and minor fonts (e.g., Calibri and Arial), and assign the updated theme back to the workbook using Aspose.Cells for C#. The example demonstrates end‑to‑end manipulation of theme data without manual editing.
// Keywords: Aspose.Cells theme XML | modify font scheme Aspose.Cells | change major font Excel .NET | update minor font in theme | reassign workbook theme C# | Excel theme manipulation Aspose | programmatic font scheme change
// Common Searches: how to change theme fonts with Aspose.Cells | extract and edit Excel theme XML in C# | set major and minor fonts in workbook theme | Aspose.Cells replace fontScheme node | apply custom theme after editing XML
// Developer Intent: Programmatically alter the major/minor fonts in an Excel theme and save the modified workbook.
// Use Cases: Enforce corporate branding by applying a standard font scheme to all generated reports. | Create reusable templates that automatically use specific fonts for headings and footnotes. | Update legacy workbooks to match a new design system without manual theme editing.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, extracts its theme XML, changes the <a:majorFont> to Calibri and <a:minorFont> to Arial, and reassigns the edited theme to the workbook. | Show a step‑by‑step example of locating the <a:fontScheme> node in the theme XML, modifying font attributes, and applying the updated theme back using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Learn how to retrieve the theme XML from an Excel workbook, modify the <a:fontScheme> node to set custom major and minor fonts (e.g., Calibri and Arial), and assign the updated theme back to the workbook using Aspose.Cells for C#. The example demonstrates end‑to‑end manipulation of theme data without manual editing.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "InputWorkbook.xlsx";
                string outputPath = "OutputWorkbook.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Show current theme name (if a theme is applied)
                Console.WriteLine("Current theme: " + (string.IsNullOrEmpty(workbook.Theme) ? "None" : workbook.Theme));

                // Set default style font to Calibri (major font) and Arial (minor font alternative)
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.Font.Name = "Calibri";
                workbook.DefaultStyle = defaultStyle;

                // Optionally, set a specific style for minor text (e.g., comments) if needed
                // Here we demonstrate creating a new style that uses Arial
                Style minorStyle = workbook.CreateStyle();
                minorStyle.Font.Name = "Arial";
                // Apply the minor style to a sample cell (A2) for illustration
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A2"];
                cell.PutValue("Minor font example");
                cell.SetStyle(minorStyle);

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
