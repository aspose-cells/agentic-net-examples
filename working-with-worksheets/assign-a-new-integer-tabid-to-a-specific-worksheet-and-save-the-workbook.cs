using System;
using Aspose.Cells;

class AssignTabIdDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Assign a new TabId value
        worksheet.TabId = 12345;

        // Save the workbook
        string filePath = "TabIdDemo.xlsx";
        workbook.Save(filePath);

        // Load the saved workbook to verify the TabId
        Workbook loadedWorkbook = new Workbook(filePath);
        int loadedTabId = loadedWorkbook.Worksheets[0].TabId;

        // Display the TabId value
        Console.WriteLine("Loaded Worksheet TabId: " + loadedTabId);
    }
}