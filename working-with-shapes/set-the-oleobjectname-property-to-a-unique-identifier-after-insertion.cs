// Title: Insert an OLE object into a worksheet cell and assign a GUID to its Name property with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to add an OLE object from a file to a specific cell, sets OleObject.Name to a newly created GUID string, and saves the workbook. | Create a reusable C# method that inserts an OLE object into a given worksheet location, assigns a unique identifier to the object's Name property, and returns the updated Workbook.
// Common Searches: Aspose.Cells C# set OleObject.Name to a GUID after inserting OLE object | how to give a unique name to an OLE object in Excel using Aspose.Cells | C# code example for adding OLE object to cell B2 with Aspose.Cells | assign unique identifier to OleObject.Name property in Aspose.Cells workbook | insert OLE object into specific worksheet cell and set Name property in Aspose.Cells
// Tags: aspose.cells insert oleobject c# | aspose.cells set oleobject name guid | aspose.cells oleobject unique identifier | aspose.cells add oleobject to worksheet cell | aspose.cells oleobject name property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# program that creates a new workbook, inserts an OLE object from 'sample.docx' into cell B2, assigns a newly generated GUID string to the OleObject.Name property, and saves the result as 'Result.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the source OLE file
            string oleFilePath = "sample.docx";

            // Ensure the source file exists; create an empty placeholder if missing
            if (!File.Exists(oleFilePath))
            {
                File.WriteAllBytes(oleFilePath, new byte[0]);
            }

            // Read the OLE file into a byte array
            byte[] oleData = File.ReadAllBytes(oleFilePath);

            // Insert an OLE object at cell B2 (row index 1, column index 1)
            // Height = 200 pixels, Width = 300 pixels, offsets set to 0
            OleObject ole = sheet.Shapes.AddOleObject(
                1,          // upperLeftRow
                1,          // upperLeftColumn
                200,        // height
                300,        // width
                0,          // upperLeftRowOffset
                0,          // upperLeftColumnOffset
                oleData     // OLE object data
            );

            // Assign a unique identifier to the OLE object's Name property
            ole.Name = Guid.NewGuid().ToString();

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
