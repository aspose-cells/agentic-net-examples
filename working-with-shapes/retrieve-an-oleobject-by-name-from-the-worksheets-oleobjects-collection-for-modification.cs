// Title: C# Example: Retrieve and Modify an OleObject by Name with Aspose.Cells
// Description: This sample loads an Excel workbook, accesses a worksheet, searches its OleObjects collection for an object whose Name matches a specified value, and updates properties such as Label and DisplayAsIcon before saving the file. It demonstrates how to locate, verify, and edit OLE objects programmatically using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | OleObject | retrieve OleObject by name | modify OleObject properties | Excel OLE object | display as icon | update OleObject label | worksheet OleObjects collection | GitHub sample | Aspose.Cells example
// Common Searches: Aspose.Cells find OleObject by name C# | change OleObject label Aspose.Cells .NET | set DisplayAsIcon for Excel OLE object using Aspose | update OleObject source file path Aspose.Cells | iterate worksheet OleObjects collection example
// Developer Intent: Locate a specific OleObject by its Name property and modify its attributes in an Excel workbook.
// Use Cases: Search for an OleObject named "MyOleObject" and change its label to "UpdatedLabel" while enabling the icon view. | Replace the source file of an existing OleObject after locating it by name. | Validate the presence of a named OleObject and log a warning if it does not exist. | Programmatically toggle the DisplayAsIcon flag for a set of OLE objects on a worksheet.
// AI Prompts: Generate C# code using Aspose.Cells that finds an OleObject called "Report" on the second worksheet and sets its ObjectSourceFullName to "C:\Data\report.xlsx". | Provide a snippet that iterates through worksheet.OleObjects, locates the object with Name "ChartOle", changes its Label to "SalesChart", and saves the workbook. | Write a reusable method that accepts a workbook path and an OleObject name, returns true if the object exists, and updates its DisplayAsIcon property to false.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This sample loads an Excel workbook, accesses a worksheet, searches its OleObjects collection for an object whose Name matches a specified value, and updates properties such as Label and DisplayAsIcon before saving the file. It demonstrates how to locate, verify, and edit OLE objects programmatically using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or use workbook.Worksheets["SheetName"] for a specific sheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Name of the OleObject we want to retrieve
        string oleObjectName = "MyOleObject";

        // Find the OleObject by its Name property
        OleObject targetOle = null;
        foreach (OleObject ole in worksheet.OleObjects)
        {
            if (ole.Name == oleObjectName)
            {
                targetOle = ole;
                break;
            }
        }

        if (targetOle != null)
        {
            // Example modification: change the label and display it as an icon
            targetOle.Label = "UpdatedLabel";
            targetOle.DisplayAsIcon = true;

            // Additional modifications can be performed here, e.g.:
            // targetOle.ObjectSourceFullName = @"C:\NewPath\file.xlsx";
        }
        else
        {
            Console.WriteLine($"OleObject with name '{oleObjectName}' was not found.");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
