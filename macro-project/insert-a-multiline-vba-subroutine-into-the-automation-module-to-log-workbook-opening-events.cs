// Title: Add a multi‑line VBA Workbook_Open subroutine to an Automation module using Aspose.Cells for .NET
// Description: Demonstrates how to create a new workbook, add a Document‑type VBA module named Automation, embed a multi‑line Workbook_Open procedure that shows a message box, and save the file as a macro‑enabled Xlsm workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA module C# | Workbook_Open event Aspose.Cells | macro‑enabled Xlsm save Aspose | add VBA code programmatically .NET | Automation module VBA Aspose.Cells
// Common Searches: add Workbook_Open handler with Aspose.Cells .NET | insert multi‑line VBA into Automation module | save workbook as Xlsm after adding VBA | create VBA Document module using Aspose.Cells
// Developer Intent: Insert a Workbook_Open VBA routine into the Automation module of a generated workbook and persist it as a macro‑enabled file.
// Use Cases: Show a custom alert each time the workbook opens. | Run initialization logic such as setting defaults or loading data on open. | Provide downstream users with built‑in VBA event handlers in programmatically created workbooks.
// AI Prompts: Write C# code with Aspose.Cells that adds a Workbook_Open subroutine which writes an entry to a log file instead of displaying a MsgBox. | Show how to add both Workbook_Open and Worksheet_Change procedures to the Automation module using Aspose.Cells. | Explain how to modify the inserted VBA to call an external COM library when the workbook is opened.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    // Demonstrates how to create a new workbook, add a Document‑type VBA module named Automation, embed a multi‑line Workbook_Open procedure that shows a message box, and save the file as a macro‑enabled Xlsm workbook with Aspose.Cells for .NET.
    public class InsertWorkbookOpenLogger
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default format is Xlsx)
                Workbook workbook = new Workbook();

                // Add a Document (Automation) module to the VBA project.
                // This module will contain the Workbook_Open event handler.
                int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Document, "Automation");

                // Retrieve the added module.
                VbaModule automationModule = workbook.VbaProject.Modules[moduleIndex];

                // VBA code that runs when the workbook is opened.
                string vbaCode =
                    "Private Sub Workbook_Open()\r\n" +
                    "    MsgBox \"Workbook has been opened.\"\r\n" +
                    "End Sub";

                // Assign the VBA code to the module.
                automationModule.Codes = vbaCode;

                // Save the workbook as a macro‑enabled file (Xlsm) so the VBA code is retained.
                string outputPath = "WorkbookWithOpenLogger.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console applications.
    public class Program
    {
        public static void Main(string[] args)
        {
            InsertWorkbookOpenLogger.Run();
        }
    }
}
