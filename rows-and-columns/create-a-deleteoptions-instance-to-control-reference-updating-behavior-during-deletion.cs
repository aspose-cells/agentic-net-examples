// Title: Delete a row while updating formula references using DeleteOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete a row and keep formulas updated | using DeleteOptions.UpdateReference to avoid #REF errors in Aspose.Cells | how to delete rows with reference updating in Aspose.Cells .NET | example of DeleteOptions with DeleteRows method in C# | preserve formula references when removing rows in Aspose.Cells
// Tags: DeleteOptions UpdateReference usage | DeleteRows with DeleteOptions | formula reference preservation Aspose.Cells | C# row removal without #REF | Aspose.Cells reference update control

using Aspose.Cells;

// Demonstrates creating a Workbook, adding sample data, configuring DeleteOptions to update formula references, deleting a specific row with DeleteRows, and saving the workbook as DeletedRowsWithOptions.xlsx.
class DeleteOptionsExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data (optional)
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Create DeleteOptions to control reference updating behavior
        DeleteOptions deleteOptions = new DeleteOptions();

        // Set UpdateReference to true to update formulas that refer to the deleted rows/columns
        // Set to false if you want the formulas to keep the original references (they may become #REF!)
        deleteOptions.UpdateReference = true;

        // Delete the second row (index 1) using the DeleteOptions instance
        sheet.Cells.DeleteRows(1, 1, deleteOptions);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("DeletedRowsWithOptions.xlsx");
    }
}
