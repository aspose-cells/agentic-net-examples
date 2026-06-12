using System;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with default worksheet
            Workbook workbook = new Workbook();

            // Add additional worksheets for demonstration
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Iterate through all worksheets and attempt to set TabId to zero
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Setting TabId to zero is expected to raise an exception
                    sheet.TabId = 0;
                    Console.WriteLine($"Worksheet \"{sheet.Name}\": TabId set to 0 successfully.");
                }
                catch (Exception ex)
                {
                    // Log the caught exception details
                    Console.WriteLine($"Worksheet \"{sheet.Name}\": Exception occurred while setting TabId to 0.");
                    Console.WriteLine($"Message: {ex.Message}");
                }
            }

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("TabIdDemoOutput.xlsx");
        }
    }
}