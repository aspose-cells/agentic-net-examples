using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

class ProtectWorksheetAfterSmartMarkers
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample smart markers
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Score");

            // Smart markers that will be replaced by data source values
            cells["A2"].PutValue("&=$Name");
            cells["B2"].PutValue("&=$Score");

            // Simple data source
            var data = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object> { ["Name"] = "Alice", ["Score"] = 85 },
                new Dictionary<string, object> { ["Name"] = "Bob",   ["Score"] = 92 }
            };

            // Bind data source and process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", data);
            designer.Process();

            // Create a style for editable cells (unlocked)
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;

            // Apply the unlocked style to the score column
            AsposeRange editableRange = worksheet.Cells.CreateRange("B2:B3");
            StyleFlag flag = new StyleFlag();
            flag.Locked = true; // Apply the Locked attribute from the style
            editableRange.ApplyStyle(unlockedStyle, flag);

            // Protect the worksheet; only unlocked cells can be edited
            // The third parameter (oldPassword) is not required here, pass null
            worksheet.Protect(ProtectionType.All, "pwd123", null);

            // Fine‑tune selection behavior
            worksheet.Protection.AllowSelectingUnlockedCell = true;
            worksheet.Protection.AllowSelectingLockedCell = false;

            // Save the workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}