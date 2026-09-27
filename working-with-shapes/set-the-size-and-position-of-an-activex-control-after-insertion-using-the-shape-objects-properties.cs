// Title: How to set Left, Top, Width, and Height of an inserted ActiveX button shape using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a button shape into a worksheet with Aspose.Cells and then adjust its Left, Top, Width, and Height properties in C#. | Update the position and size of an ActiveX control after adding it to an Excel sheet using the Shape object's properties in Aspose.Cells for .NET. | Programmatically move and resize a form control button in an Excel workbook using the Aspose.Cells Shape API in C#.
// Common Searches: Aspose.Cells C# set button shape left and top coordinates | Change size of ActiveX button inserted with Aspose.Cells .NET | Move Excel form control after adding with Aspose.Cells API | Adjust shape dimensions programmatically using Aspose.Cells in C# | Set width and height of a Shape object in Aspose.Cells workbook
// Tags: Aspose.Cells shape positioning C# | set button shape dimensions Aspose.Cells | adjust ActiveX control size .NET | modify shape left top properties Excel | Aspose.Cells insert button shape

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Demonstrates creating a new workbook, adding a button (form control) shape, and using the Shape object's Left, Top, Width, and Height properties to position and size the ActiveX control before saving the workbook with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Insert a button (form control) into the sheet
                // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
                Shape buttonShape = sheet.Shapes.AddButton(2, 1, 4, 3, 200, 30);

                // Set the position (in points) of the button
                buttonShape.Left = 100;   // distance from the left edge of the worksheet
                buttonShape.Top = 50;     // distance from the top edge of the worksheet

                // Optional: set button caption
                buttonShape.Text = "Click Me";

                // Save the workbook to a file
                string outputPath = "ActiveXControl.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
