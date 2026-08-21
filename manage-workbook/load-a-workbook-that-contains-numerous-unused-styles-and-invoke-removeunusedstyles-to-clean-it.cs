// Title: C# – Remove Unused Styles from an Excel Workbook with Aspose.Cells (Workbook.RemoveUnusedStyles)
// Description: Demonstrates how to load an Excel workbook (or create a sample when missing), display the style pool count, invoke Workbook.RemoveUnusedStyles to purge unused cell styles, and save the optimized file. Includes error handling and resource cleanup.
// Keywords: Aspose.Cells | Workbook.RemoveUnusedStyles | C# Excel style cleanup | CountOfStylesInPool | remove unused cell styles | .NET Excel optimization | clean Excel style pool | save cleaned workbook | sample workbook creation
// Common Searches: Aspose.Cells remove unused styles C# | Workbook.RemoveUnusedStyles example | how to count Excel styles before removal | delete unused cell styles with Aspose.Cells | create sample workbook when file not found Aspose.Cells
// Developer Intent: Purge all unused cell styles from a workbook and save the reduced‑size file.
// Use Cases: Reduce file size of large Excel reports before distribution. | Automate style‑pool cleanup in server‑side Excel generation pipelines. | Validate style usage by logging counts before and after removal. | Provide a fallback sample workbook when the source file is unavailable.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, shows the style count, calls RemoveUnusedStyles, and saves the result. | Write a method that checks for a workbook file, creates a styled sample workbook if absent, cleans unused styles, and returns the cleaned workbook object. | Explain how Workbook.CountOfStylesInPool and Workbook.RemoveUnusedStyles work together to minimize Excel file size.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to load an Excel workbook (or create a sample when missing), display the style pool count, invoke Workbook.RemoveUnusedStyles to purge unused cell styles, and save the optimized file. Includes error handling and resource cleanup.
    public class RemoveUnusedStylesDemo
    {
        public static void Main()
        {
            // Path to the source workbook that contains many unused styles
            string sourcePath = "InputWorkbookWithUnusedStyles.xlsx";

            Workbook workbook = null;

            try
            {
                // Load the workbook if the file exists; otherwise create a sample workbook
                if (File.Exists(sourcePath))
                {
                    workbook = new Workbook(sourcePath);
                }
                else
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    Console.WriteLine("Creating a sample workbook for demonstration.");

                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Sample";

                    // Add some styled cells to generate a style pool
                    Style style = workbook.CreateStyle();
                    style.Font.Color = System.Drawing.Color.Blue;
                    style.Font.IsBold = true;

                    StyleFlag flag = new StyleFlag { FontColor = true, FontBold = true };

                    Cell cell = sheet.Cells["A1"];
                    cell.PutValue("Styled Text");
                    cell.SetStyle(style, flag);
                }

                // Display the number of styles in the style pool before cleanup
                Console.WriteLine("Styles before removal: " + workbook.CountOfStylesInPool);

                // Remove all unused styles from the workbook
                workbook.RemoveUnusedStyles();

                // Display the number of styles after cleanup
                Console.WriteLine("Styles after removal: " + workbook.CountOfStylesInPool);

                // Save the cleaned workbook to a new file
                string outputPath = "CleanedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved after removing unused styles: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                // Ensure resources are released
                if (workbook != null)
                {
                    workbook.Dispose();
                }
            }
        }
    }
}
