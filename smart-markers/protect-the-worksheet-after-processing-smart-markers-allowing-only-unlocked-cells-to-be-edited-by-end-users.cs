// Title: Protect an Excel worksheet after processing smart markers while keeping specific cells unlocked using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a workbook, processes smart markers with WorkbookDesigner, unlocks a specified cell range, and then applies full worksheet protection while preserving the unlocked cells. | Show how to apply an unlocked style to a range and configure worksheet protection options (allow selecting locked and unlocked cells) after smart marker processing with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells protect worksheet after smart marker processing C# example | unlock cells B2:B10 before protecting sheet with Aspose.Cells | keep certain cells editable when protecting Excel sheet using Aspose.Cells .NET | WorkbookDesigner smart markers unlock range then protect worksheet
// Tags: Aspose.Cells worksheet protection with unlocked cells | smart markers processing unlock range C# | WorkbookDesigner apply style flag locked property | Excel sheet protect allow selecting unlocked cells Aspose | C# unlock cell range before worksheet protection Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

// The sample loads Input.xlsx, creates a DataSet with employee data, processes smart markers via WorkbookDesigner, unlocks cells B2:B10, protects the first worksheet with all protection types while allowing selection of both locked and unlocked cells, and saves the result to Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains smart markers
            var workbook = new Workbook(inputPath);

            // Prepare a DataSet as the data source for the smart markers
            var dataSet = new DataSet();

            var employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));

            employeeTable.Rows.Add("John", 30);
            employeeTable.Rows.Add("Jane", 25);

            dataSet.Tables.Add(employeeTable);

            // Process the smart markers using the data source
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Unlock cells that should remain editable by users
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // Example: unlock the range B2:B10 (adjust as needed)
            var unlockedRange = cells.CreateRange("B2:B10");
            var unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false; // mark cells as unlocked
            var styleFlag = new StyleFlag { Locked = true }; // apply only the lock flag
            unlockedRange.ApplyStyle(unlockedStyle, styleFlag);

            // Protect the worksheet while keeping unlocked cells editable
            sheet.Protect(ProtectionType.All);
            sheet.Protection.AllowSelectingLockedCell = true;
            sheet.Protection.AllowSelectingUnlockedCell = true;

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
