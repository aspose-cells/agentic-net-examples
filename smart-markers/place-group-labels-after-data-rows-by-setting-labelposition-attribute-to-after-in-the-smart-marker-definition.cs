using System;
using System.IO;
using Aspose.Cells;

class GroupLabelsAfterDataRows
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Item 3");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("Item 4");
            sheet.Cells["B5"].PutValue(40);

            // Insert a smart marker that groups rows and places the group label after the data rows
            // The attribute LabelPosition=After tells Aspose.Cells to put the summary label below the detail rows
            sheet.Cells["C1"].PutValue("&=Group(A2:A5,LabelPosition=After)");

            // Process the smart markers using WorkbookDesigner (the correct API for smart markers)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.Process();

            // Ensure that the outline (group) summary row is positioned below the detail rows
            sheet.Outline.SummaryRowBelow = true;

            // Define output file
            string outputPath = "GroupLabelsAfterDataRows.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}