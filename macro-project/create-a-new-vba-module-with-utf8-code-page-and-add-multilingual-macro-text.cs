// Title: Add a UTF‑8 VBA class module with multilingual messages using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create an empty workbook, access its VbaProject, set the project name and UTF‑8 encoding, add a class module named MultiLangModule, embed VBA code that shows MsgBox messages in English, Chinese, and Arabic, and save the workbook as a macro‑enabled XLSM file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA module C# | UTF-8 VBA project encoding | multilingual VBA macro | macro‑enabled XLSM workbook | add class module Aspose.Cells | international Excel VBA | C# Excel automation | Aspose.Cells for .NET examples
// Common Searches: set VBA project encoding to UTF-8 with Aspose.Cells | add a class module to a VBA project in C# | create multilingual VBA macro in an XLSM file | Aspose.Cells example for VBA modules | save workbook with VBA macros using Aspose.Cells
// Developer Intent: Create a VBA class module encoded in UTF‑8, insert a macro that displays messages in several languages, and export the workbook as a macro‑enabled XLSM file.
// Use Cases: Generate Excel templates that contain localized VBA alerts for global users. | Automate distribution of macro‑enabled reports with language‑specific instructions. | Build prototype workbooks that demonstrate internationalization of VBA code.
// AI Prompts: Write C# code with Aspose.Cells to add a VBA class module, set UTF‑8 encoding, and include a macro that shows messages in English, Chinese, and Arabic. | Explain how to change the VBA project name and verify the encoding after creating a workbook with Aspose.Cells. | Provide a step‑by‑step guide to test that multilingual VBA code is correctly stored in a class module before saving the XLSM file.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaMultilingualDemo
{
    // Demonstrates how to create an empty workbook, access its VbaProject, set the project name and UTF‑8 encoding, add a class module named MultiLangModule, embed VBA code that shows MsgBox messages in English, Chinese, and Arabic, and save the workbook as a macro‑enabled XLSM file with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Set the VBA project name (optional) and encoding to UTF‑8
            vbaProject.Name = "MultilingualVbaProject";
            vbaProject.Encoding = Encoding.UTF8;

            // Add a new class module to the VBA project
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MultiLangModule");
            VbaModule vbaModule = vbaProject.Modules[moduleIndex];

            // Define multilingual VBA macro code
            string macroCode = 
                "Sub MultiLangMacro()\r\n" +
                "    MsgBox \"Hello\"          ' English\r\n" +
                "    MsgBox \"你好\"          ' Chinese\r\n" +
                "    MsgBox \"مرحبا\"        ' Arabic\r\n" +
                "End Sub";

            // Assign the code to the module
            vbaModule.Codes = macroCode;

            // Save the workbook as a macro‑enabled file
            workbook.Save("MultilingualVbaDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}
