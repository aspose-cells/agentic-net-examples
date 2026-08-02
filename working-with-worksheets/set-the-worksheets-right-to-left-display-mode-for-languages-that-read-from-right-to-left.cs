// Title: Set Worksheet Right‑to‑Left Display Mode with Aspose.Cells for .NET
// Description: Demonstrates how to enable right‑to‑left (RTL) orientation for an Excel worksheet using Aspose.Cells for .NET by setting the DisplayRightToLeft property, saving the workbook, reloading it, and confirming the setting.
// Keywords: Aspose.Cells | .NET | C# | DisplayRightToLeft | right to left | RTL mode | worksheet orientation | Excel RTL | Arabic Excel | Hebrew Excel | save workbook | load workbook | verify RTL flag
// Common Searches: Aspose.Cells set worksheet RTL | Enable right‑to‑left display in Excel using C# | DisplayRightToLeft property example | How to save Excel file with RTL layout in .NET | Check RTL setting after workbook reload
// Developer Intent: Enable RTL layout for a worksheet and ensure the setting persists after saving.
// Use Cases: Create Excel reports for Arabic or Hebrew audiences where the sheet must read from right to left. | Generate templates with RTL orientation before populating data programmatically. | Validate that the RTL flag remains intact when the file is opened in other spreadsheet applications.
// AI Prompts: Show code to set DisplayRightToLeft for all worksheets in a workbook using Aspose.Cells. | Provide a C# example that toggles RTL mode based on the system culture (e.g., ar‑SA, he‑IL). | Explain how to read the DisplayRightToLeft flag from an existing Excel file with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable right‑to‑left (RTL) orientation for an Excel worksheet using Aspose.Cells for .NET by setting the DisplayRightToLeft property, saving the workbook, reloading it, and confirming the setting.
    public class SetWorksheetRightToLeftDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable right‑to‑left display mode for the worksheet
                worksheet.DisplayRightToLeft = true;

                // Save the workbook to a file
                string filePath = "RightToLeftDemo.xlsx";
                workbook.Save(filePath);

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Reload the workbook to confirm the setting
                Workbook loadedWorkbook = new Workbook(filePath);
                Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

                // Output the current DisplayRightToLeft value
                Console.WriteLine("DisplayRightToLeft: " + loadedWorksheet.DisplayRightToLeft);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetWorksheetRightToLeftDemo.Run();
        }
    }
}
