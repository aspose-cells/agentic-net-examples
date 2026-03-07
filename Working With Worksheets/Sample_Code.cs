using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsSample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the code name of the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.CodeName = "MainSheet";

            // Add a VBA class module and set its code
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "MyClass");
            VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];
            vbaModule.Codes = "Public Sub Hello()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

            // Retrieve binary codes of the module (for demonstration)
            byte[] binary = vbaModule.BinaryCodes;
            Console.WriteLine("Binary codes length: " + binary.Length);
            // Optionally convert to a string (may contain non‑text data)
            string binaryString = Encoding.UTF8.GetString(binary);
            Console.WriteLine("Binary codes (as string snippet): " + binaryString.Substring(0, Math.Min(50, binaryString.Length)));

            // Demonstrate handling of CellsException
            try
            {
                // Attempt to access a non‑existent worksheet
                Worksheet missing = workbook.Worksheets["NonExistent"];
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Caught CellsException. Code: {ex.Code}, Message: {ex.Message}");
            }

            // Save the workbook as a macro‑enabled file
            workbook.Save("SampleDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}