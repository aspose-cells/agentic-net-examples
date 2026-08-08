// Title: C# – Protect Worksheet After Smart Marker Processing While Unlocking Specific Cells (Aspose.Cells)
// Description: Loads a template workbook with smart markers, fills it using a List<Person>, unlocks the range B2:C10, applies worksheet protection with a password, and saves the protected file. Demonstrates how to combine smart marker processing with selective cell editing in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Smart markers | worksheet protection | unlock cells | IsLocked style | password protection | Excel automation | range protection
// Common Searches: Aspose.Cells protect worksheet after smart markers C# | unlock cells before sheet protection Aspose.Cells | set IsLocked false for a range Aspose.Cells | apply password to worksheet while allowing edits | smart markers worksheet lock .NET
// Developer Intent: Lock a worksheet after populating smart markers, keeping only chosen cells editable for end users.
// Use Cases: Generate a sales dashboard from a template, then lock the sheet so users can only modify the input range B2:C10. | Create a collaborative timesheet where employee data is filled via smart markers and only comment cells remain editable. | Automate invoice generation, protect the invoice sheet, and allow editing of the payment‑status column while all other fields stay read‑only.
// AI Prompts: Show C# code that protects an Aspose.Cells worksheet after processing smart markers, unlocking a specific range. | How do I apply an unlocked style to a cell range and then secure the sheet with a password in Aspose.Cells? | Explain the role of StyleFlag when setting IsLocked = false before worksheet protection in Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerProtection
{
    // Loads a template workbook with smart markers, fills it using a List<Person>, unlocks the range B2:C10, applies worksheet protection with a password, and saves the protected file. Demonstrates how to combine smart marker processing with selective cell editing in Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                const string templatePath = "template.xlsx";
                const string outputPath = "output_protected.xlsx";

                // Ensure the template file exists before loading
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template file not found: {templatePath}");

                // Load a workbook that contains smart markers
                Workbook workbook = new Workbook(templatePath);

                // Prepare a simple data source
                var data = new List<Person>
                {
                    new Person { Name = "John Doe", Age = 30 },
                    new Person { Name = "Jane Smith", Age = 28 }
                };

                // Set the data source for the smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("People", data);
                designer.Process();

                // Unlock the cells that should be editable by end users (e.g., B2:C10)
                Worksheet sheet = workbook.Worksheets[0];
                Aspose.Cells.Range editableRange = sheet.Cells.CreateRange("B2:C10");
                Style unlockedStyle = workbook.CreateStyle();
                unlockedStyle.IsLocked = false; // make cells editable when sheet is protected
                StyleFlag flag = new StyleFlag { Locked = true };
                editableRange.ApplyStyle(unlockedStyle, flag);

                // Protect the worksheet so that only unlocked cells can be edited
                sheet.Protect(ProtectionType.All, "securePassword123", null);

                // Save the resulting workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Simple POCO class used as a data source for smart markers
        public class Person
        {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
        }
    }
}
