// Title: Display all worksheet tabs, ensure scrollbars are visible, and save a workbook as XLSX with Aspose.Cells for .NET
// Common Searches: how to enable worksheet tabs in a new workbook using Aspose.Cells for .NET | Aspose.Cells C# example to display sheet tabs and export to XLSX | what is the default scroll bar visibility in Aspose.Cells workbooks | saving a workbook with visible tabs using the Aspose.Cells API | C# code to set ShowTabs property and write Result.xlsx
// Tags: Aspose.Cells ShowTabs property | show worksheet tabs C# | export workbook to XLSX Aspose.Cells | default scroll bar visibility Aspose.Cells | Workbook.Settings configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new Workbook, enables the Settings.ShowTabs flag so all worksheet tabs are displayed (scrollbars are shown by default), and saves the file as Result.xlsx in XLSX format, while handling any exceptions and outputting the full path.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Show all worksheet tabs
                workbook.Settings.ShowTabs = true;

                // The current Aspose.Cells API does not expose scroll‑bar visibility properties.
                // Default behavior shows scroll bars, so no additional settings are required.

                // Define output file path
                string outputPath = "Result.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
