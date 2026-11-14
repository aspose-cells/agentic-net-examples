// Title: How to apply separate open and modify passwords to an XLS workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set different passwords for opening and editing an XLS file | protect Excel .xls workbook with distinct open and write protection passwords using Aspose.Cells | C# code to assign workbook.Settings.Password and WriteProtection.Password in Aspose.Cells | set read‑only password and modify password on XLS workbook with Aspose.Cells .NET | example of applying both opening and modify passwords to an Excel file in C#
// Tags: Aspose.Cells set open password XLS | Aspose.Cells write protection password C# | Workbook.Settings.Password API | Workbook.WriteProtection.Password usage | protect XLS workbook with distinct passwords

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, writes sample data, assigns an opening password via workbook.Settings.Password, assigns a modify password via workbook.Settings.WriteProtection.Password, and saves the file as ProtectedWorkbook.xls, demonstrating how to protect an XLS workbook with separate credentials for opening and editing.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Add sample data to the first worksheet
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");

                // Set the password required to open the file
                workbook.Settings.Password = "OpenPassword123";

                // Set the password required to modify the file
                workbook.Settings.WriteProtection.Password = "ModifyPassword456";

                // Save the workbook; format is inferred from the file extension
                workbook.Save("ProtectedWorkbook.xls");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
