// Title: C# – Create a UTF‑8 VBA class module with multilingual MsgBox macros using Aspose.Cells
// Description: Demonstrates how to generate an in‑memory workbook, set the VBA project's Encoding to UTF‑8, add a class module containing MsgBox greetings in English, Chinese, Arabic, Russian and Japanese, and save it as a macro‑enabled XLSM file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA encoding UTF-8 | C# add VBA class module | multilingual VBA macro | MsgBox localization Excel | save workbook as .xlsm | programmatic VBA injection .NET | Excel macro-enabled file | Unicode VBA project Aspose | globalized Excel messages | Aspose.Cells VbaProject Encoding
// Common Searches: set VBA project encoding to UTF-8 using Aspose.Cells | add a class module with multilingual code in C# | create macro‑enabled XLSX with Aspose.Cells | how to inject VBA macros programmatically .NET | Unicode MsgBox examples in Excel VBA
// Developer Intent: Programmatically build a macro‑enabled workbook that contains a UTF‑8 encoded VBA class module with language‑specific MsgBox statements.
// Use Cases: Generate a template that shows localized greetings when the macro runs. | Inject a reusable multilingual VBA class into existing XLSM files during automated report creation. | Produce localized Excel dashboards that display language‑appropriate messages without manual editing.
// AI Prompts: Write C# code with Aspose.Cells to create a VBA class module, set its Encoding to UTF‑8, add MsgBox statements in English, Chinese, Arabic, Russian, and Japanese, and save the workbook as .xlsm. | Explain step‑by‑step how to change a VBA project's encoding to UTF‑8 and add a multilingual class module using Aspose.Cells for .NET.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaMultilingualDemo
{
    // Demonstrates how to generate an in‑memory workbook, set the VBA project's Encoding to UTF‑8, add a class module containing MsgBox greetings in English, Chinese, Arabic, Russian and Japanese, and save it as a macro‑enabled XLSM file with Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (in memory)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Set the VBA project encoding to UTF‑8 to support multilingual characters
            vbaProject.Encoding = Encoding.UTF8;

            // Add a new class module to the VBA project
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MultiLangModule");
            VbaModule module = vbaProject.Modules[moduleIndex];

            // Define multilingual VBA macro code
            // English, Chinese, Arabic, Russian, Japanese examples
            string macroCode = 
                "Sub ShowMultilingualMessage()\r\n" +
                "    ' English\r\n" +
                "    MsgBox \"Hello, World!\"\r\n" +
                "\r\n" +
                "    ' 中文 (Chinese)\r\n" +
                "    MsgBox \"你好，世界！\"\r\n" +
                "\r\n" +
                "    ' العربية (Arabic)\r\n" +
                "    MsgBox \"مرحبا بالعالم!\"\r\n" +
                "\r\n" +
                "    ' Русский (Russian)\r\n" +
                "    MsgBox \"Привет, мир!\"\r\n" +
                "\r\n" +
                "    ' 日本語 (Japanese)\r\n" +
                "    MsgBox \"こんにちは、世界！\"\r\n" +
                "End Sub";

            // Assign the macro code to the module
            module.Codes = macroCode;

            // Save the workbook as a macro‑enabled file
            workbook.Save("MultilingualVbaModule.xlsm", SaveFormat.Xlsm);
        }
    }
}
