// Title: Select a worksheet by name, create it if missing, and obtain its Cells collection using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing Excel file or creates a new workbook, selects a worksheet by its name (adding the sheet when it does not exist), and returns the Worksheet.Cells object. | Show how to read the current value of cell A1 and then write a new string to that cell after acquiring the Cells collection from a named worksheet with Aspose.Cells.
// Common Searches: Aspose.Cells C# get worksheet by name and add it if not present | C# retrieve Cells collection from a specific sheet using Aspose.Cells | How to read and write a cell after selecting a worksheet by name in Aspose.Cells .NET | Load existing Excel file or create new one then access cells of a named worksheet Aspose.Cells | Aspose.Cells example for checking worksheet existence and accessing cell A1
// Tags: worksheet name selection Aspose.Cells | auto-create missing sheet Aspose.Cells | cells collection access Aspose.Cells | cell A1 read write Aspose.Cells | excel workbook initialization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook or creates a new one, ensures a worksheet named "MyWorksheet" exists (adding it if necessary), obtains its Cells collection, reads the value of cell A1, writes "Hello Aspose!" to A1, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string filePath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if present; otherwise create a new one.
                if (File.Exists(filePath))
                {
                    workbook = new Workbook(filePath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Save(filePath);
                }

                // Select worksheet by name; create it if it does not exist.
                string sheetName = "MyWorksheet";
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    int idx = workbook.Worksheets.Add();
                    sheet = workbook.Worksheets[idx];
                    sheet.Name = sheetName;
                }

                // Obtain the Cells collection.
                Cells cells = sheet.Cells;

                // Read value from A1 (empty if not set).
                string value = cells["A1"].StringValue;
                Console.WriteLine($"Value in A1: '{value}'");

                // Write a value to A1 and save the workbook.
                cells["A1"].PutValue("Hello Aspose!");
                workbook.Save(filePath);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
