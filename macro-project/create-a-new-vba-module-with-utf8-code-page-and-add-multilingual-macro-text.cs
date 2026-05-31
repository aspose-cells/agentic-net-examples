using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CreateVbaModuleWithMultilingualMacro
{
    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Set the VBA project encoding to UTF‑8
        vbaProject.Encoding = Encoding.UTF8;

        // Add a class module named "MultiLangModule"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MultiLangModule");

        // Retrieve the added module
        VbaModule module = vbaProject.Modules[moduleIndex];

        // Multilingual VBA macro (English, Chinese, Arabic)
        string vbaCode =
            "Sub ShowMessages()\r\n" +
            "    MsgBox \"Hello World!\"  ' English\r\n" +
            "    MsgBox \"你好，世界！\"   ' Chinese\r\n" +
            "    MsgBox \"مرحبا بالعالم!\" ' Arabic\r\n" +
            "End Sub";

        // Assign the code to the module
        module.Codes = vbaCode;

        // Save the workbook as a macro‑enabled file
        workbook.Save("MultilingualVbaModule.xlsm", SaveFormat.Xlsm);
    }
}

class Program
{
    static void Main()
    {
        try
        {
            CreateVbaModuleWithMultilingualMacro.Run();
            Console.WriteLine("Workbook created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}