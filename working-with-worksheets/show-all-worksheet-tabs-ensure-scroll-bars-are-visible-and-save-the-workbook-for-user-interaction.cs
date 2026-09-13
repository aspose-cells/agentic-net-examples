// Title: How to display worksheet tabs, maintain default scrollbars, and save an Excel workbook with Aspose.Cells in C#
// AI Prompts: Generate C# code that sets Workbook.Settings.ShowTabs to true, adds data to the first worksheet, and saves the workbook as an .xlsx file using Aspose.Cells. | Write a .NET snippet that ensures worksheet tabs are visible, confirms scrollbars remain enabled, and exports the workbook for user interaction with Aspose.Cells. | Create an example that creates a new workbook, makes sheet tabs visible, inserts sample data, and writes the file to disk with Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# show worksheet tabs in generated Excel file | How to keep scrollbars visible when creating a workbook with Aspose.Cells .NET | Save Excel workbook with visible sheet tabs using Aspose.Cells in C# | C# Aspose.Cells Workbook.Settings.ShowTabs example | Export interactive Excel file with Aspose.Cells in .NET
// Tags: Workbook.Settings.ShowTabs usage | display worksheet tabs Aspose.Cells | export Excel as Xlsx using Aspose.Cells | C# ensure scrollbars visible Aspose.Cells | interactive Excel file generation .NET

using System;
using System.IO;
using Aspose.Cells;

// // Demonstrates creating a new workbook, enabling worksheet tabs, adding sample data, and saving the file as Result.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one with new Workbook("input.xlsx"))
            Workbook workbook = new Workbook();

            // Ensure worksheet tabs are visible
            workbook.Settings.ShowTabs = true;

            // Scroll bars are shown by default; explicit properties are not available in this API version.

            // Add sample data to demonstrate the workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Save the workbook for user interaction
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
