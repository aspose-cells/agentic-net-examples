// Title: How to delete a specific slicer by name from an Aspose.Cells workbook using C#
// AI Prompts: Write C# code with Aspose.Cells that locates a slicer named "MySlicer" in a worksheet, removes it from the slicers collection, and saves the workbook. | Show the steps to retrieve a slicer via its Name property from a worksheet's Slicers collection and invoke the Remove method in Aspose.Cells for .NET. | Provide a complete example that creates a table, adds a slicer, deletes the slicer by its identifier, and saves the file using the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# delete slicer using Name property | remove a slicer from an Excel workbook with Aspose.Cells .NET | how to programmatically delete a specific slicer in Aspose.Cells | C# code to find and remove slicer named MySlicer in a worksheet | clean up unused slicers in Aspose.Cells workbook
// Tags: Aspose.Cells slicer removal C# | C# Aspose.Cells slicer removal by identifier | Excel slicer cleanup Aspose.Cells | Aspose.Cells worksheet slicer management | slicer deletion from workbook .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

// The sample creates a workbook, adds a table and a slicer named "MySlicer", retrieves that slicer by its Name property, removes it from the worksheet's Slicers collection, and saves the file as DeletedSlicer.xlsx.
class DeleteSlicerByName
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data for a table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(30);

        // Add a ListObject (table) covering the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Add a slicer linked to the first column of the table and give it a name
        int slicerIndex = worksheet.Slicers.Add(table, 0, "C1");
        Slicer slicer = worksheet.Slicers[slicerIndex];
        slicer.Name = "MySlicer";

        // Retrieve the slicer by its name and remove it from the collection
        Slicer slicerToRemove = worksheet.Slicers["MySlicer"];
        worksheet.Slicers.Remove(slicerToRemove);

        // Save the workbook to verify that the slicer has been removed
        workbook.Save("DeletedSlicer.xlsx");
    }
}
