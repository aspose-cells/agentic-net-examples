// Title: Set Worksheet Right‑to‑Left Display Mode with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable the right‑to‑left (RTL) view for an Excel worksheet using Aspose.Cells for .NET. The example creates a workbook, sets the Worksheet.DisplayRightToLeft property, saves the file, reloads it, and confirms the setting—ideal for Arabic, Hebrew, or other RTL language users.
// Keywords: Aspose.Cells C# | Worksheet DisplayRightToLeft | right to left Excel .NET | RTL worksheet Aspose | Arabic Excel display mode | Hebrew Excel layout | set RTL mode programmatically | Excel right‑to‑left orientation
// Common Searches: Aspose.Cells enable right to left view C# | DisplayRightToLeft property example | How to set RTL mode for Excel sheet using Aspose | C# code for Arabic Excel layout with Aspose.Cells | Toggle worksheet right‑to‑left in .NET
// Developer Intent: Programmatically switch a worksheet’s layout to right‑to‑left so that content aligns correctly for RTL languages such as Arabic or Hebrew.
// Use Cases: Create a new workbook and activate RTL display for the first sheet before distributing to Middle‑East users. | Open an existing Excel file, set DisplayRightToLeft on a specific worksheet, and verify the setting after saving. | Apply RTL mode to every worksheet in a workbook to prepare a multilingual report for Arabic and Hebrew audiences.
// AI Prompts: Write C# code that sets Worksheet.DisplayRightToLeft = true for a given sheet in an Aspose.Cells workbook. | Provide a loop that iterates through all worksheets in a workbook and enables right‑to‑left display, then saves the file. | Explain how to confirm that the RTL setting persists after reloading the workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable the right‑to‑left (RTL) view for an Excel worksheet using Aspose.Cells for .NET. The example creates a workbook, sets the Worksheet.DisplayRightToLeft property, saves the file, reloads it, and confirms the setting—ideal for Arabic, Hebrew, or other RTL language users.
    public class WorksheetDisplayRightToLeftDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable right-to-left display mode for the worksheet
                worksheet.DisplayRightToLeft = true;

                // Save the workbook
                string filePath = "RightToLeftDemo.xlsx";
                workbook.Save(filePath);

                // Verify the setting by loading the saved workbook
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
                    Console.WriteLine("DisplayRightToLeft: " + loadedWorksheet.DisplayRightToLeft);
                }
                else
                {
                    Console.WriteLine($"Error: File '{filePath}' was not found after saving.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetDisplayRightToLeftDemo.Run();
        }
    }
}
