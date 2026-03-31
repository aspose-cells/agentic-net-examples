using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetInsertion
{
    class Program
    {
        static void Main()
        {
            // Initialize a WorkbookDesigner instance (lifecycle rule)
            WorkbookDesigner designer = new WorkbookDesigner();

            // Load an existing template workbook (lifecycle rule)
            // Replace "Template.xlsx" with the path to your designer spreadsheet
            designer.Workbook = new Workbook("Template.xlsx");

            // Insert a new worksheet at index 1 with default worksheet type (Insert rule)
            Worksheet insertedSheet = designer.Workbook.Worksheets.Insert(1, SheetType.Worksheet);
            insertedSheet.Name = "InsertedSheet";
            // Example: add some data to the inserted sheet
            insertedSheet.Cells["A1"].PutValue("Data in inserted sheet");

            // Add another worksheet by name (Add rule)
            Worksheet addedSheet = designer.Workbook.Worksheets.Add("AddedSheet");
            addedSheet.Cells["A1"].PutValue("Data in added sheet");

            // If the template contains smart markers, process them so that new data integrates correctly
            designer.Process();

            // Save the modified workbook (lifecycle rule)
            // Replace "Result.xlsx" with your desired output path
            designer.Workbook.Save("Result.xlsx");
        }
    }
}