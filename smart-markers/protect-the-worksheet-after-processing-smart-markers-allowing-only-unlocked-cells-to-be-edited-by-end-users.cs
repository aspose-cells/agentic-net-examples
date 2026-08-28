// Title: How to process smart markers, unlock specific cells, and protect an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a workbook containing smart markers, fills it using WorkbookDesigner, unlocks a given cell range, and then applies worksheet protection with a password using Aspose.Cells. | Show the steps to create an unlocked style, assign it to a range after smart marker processing, and protect the sheet so only those cells remain editable.
// Common Searches: Aspose.Cells C# unlock cells before protecting worksheet after smart marker processing | How to keep certain Excel cells editable when protecting a sheet with Aspose.Cells | Protect Excel sheet with password while allowing B2:B3 range to be edited using Aspose.Cells | Smart markers processing then worksheet protection example in C# | Set IsLocked false for cells after WorkbookDesigner.Process in Aspose.Cells
// Tags: process smart markers Aspose.Cells C# | make cells editable before sheet protection Aspose.Cells | apply unlocked style to range Aspose.Cells | protect worksheet with password Aspose.Cells | WorkbookDesigner data source example Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This C# example loads a template workbook containing smart markers, populates them with a List<Person> via WorkbookDesigner, unlocks the B2:B3 cells by applying an unlocked style, protects the first worksheet with a password, and saves the resulting protected file.
public class SmartMarkerProtectionExample
{
    public static void Main()
    {
        try
        {
            const string templatePath = "TemplateWithSmartMarkers.xlsx";
            const string resultPath = "ResultProtected.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load a workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Prepare a simple data source for the smart markers
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 28 }
            };

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Persons", persons);

            // Process all smart markers in the workbook
            designer.Process();

            // Unlock the cells that should remain editable after protection
            // Example: unlock the range B2:B3 (where ages will be placed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            AsposeRange editableRange = cells.CreateRange("B2:B3");

            // Create a style with IsLocked = false
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;

            // Apply the unlocked style to each cell in the range
            int firstRow = editableRange.FirstRow;
            int firstColumn = editableRange.FirstColumn;
            int rowCount = editableRange.RowCount;
            int columnCount = editableRange.ColumnCount;

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = cells[firstRow + i, firstColumn + j];
                    cell.SetStyle(unlockedStyle);
                }
            }

            // Protect the worksheet; only unlocked cells can be edited by end users
            // Provide an empty oldPassword as required by the API overload
            sheet.Protect(ProtectionType.All, "pwd123", string.Empty);

            // Save the processed and protected workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple POCO class used as a data source for smart markers
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
