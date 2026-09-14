// Title: How to protect workbook windows (prevent zoom level changes) with Aspose.Cells for .NET
// AI Prompts: Generate C# code that applies window protection to an Excel workbook, optionally setting a password, using Aspose.Cells. | Illustrate how to lock only the view‑related settings of an existing spreadsheet and save the protected file. | Provide an example of enforcing workbook window protection without supplying a password via Aspose.Cells.
// Common Searches: Aspose.Cells protect workbook windows to lock zoom level in C# | disable default view changes in an Excel file using .NET | C# protect workbook view settings without password Aspose.Cells
// Tags: Aspose.Cells Windows protection API | prevent Excel zoom level changes C# | protect workbook view settings Aspose.Cells | disable workbook window modifications .NET | Excel workbook window protection without password

using System;
using Aspose.Cells;

namespace WorkbookProtectionExampleApp
{
    // The example creates (or loads) an Excel workbook, applies workbook window protection to block changes to default view settings such as zoom level—optionally with a password—and saves the result as "ProtectedWorkbook.xlsx".
    class WorkbookProtectionExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // Use new Workbook("input.xlsx") to load an existing file

                // Protect the workbook windows (default view settings such as zoom level) with an optional password
                workbook.Protect(ProtectionType.Windows, ""); // Empty password disables password protection

                // Save the protected workbook
                string outputPath = "ProtectedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
