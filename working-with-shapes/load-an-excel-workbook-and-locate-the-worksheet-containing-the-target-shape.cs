// Title: Find the worksheet that contains a specific shape name in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets and returns the name of the sheet where a shape with a given Name is found. | Show how to exit nested loops efficiently after locating a target shape in a workbook and optionally save the modified file.
// Common Searches: asp.net locate worksheet that holds a shape named MyShape using Aspose.Cells | c# iterate over workbook worksheets to find a specific drawing shape | asp.net core retrieve sheet index of a shape in an Excel file with Aspose.Cells | how to determine which worksheet contains a particular shape in Excel via Aspose.Cells | c# break out of nested loops after finding a shape in Aspose.Cells workbook
// Tags: shape name lookup Aspose.Cells | worksheet identification from shape C# | loop through workbook shapes Aspose.Cells | conditional break after shape detection .NET

// Load the Excel workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Define the name (or other identifier) of the target shape
string targetShapeName = "MyShape";

// Variable to hold the worksheet that contains the shape
Aspose.Cells.Worksheet targetWorksheet = null;

// Iterate through all worksheets in the workbook
foreach (Aspose.Cells.Worksheet sheet in workbook.Worksheets)
{
    // Check if the worksheet has any shapes
    if (sheet.Shapes.Count > 0)
    {
        // Iterate through the shapes in the current worksheet
        foreach (Aspose.Cells.Drawing.Shape shape in sheet.Shapes)
        {
            // Compare shape name (or other property) with the target identifier
            if (shape.Name == targetShapeName)
            {
                targetWorksheet = sheet;
                break;
            }
        }
    }

    // If the shape has been found, exit the outer loop as well
    if (targetWorksheet != null)
        break;
}

// At this point, targetWorksheet holds the worksheet containing the shape (or null if not found)
if (targetWorksheet != null)
{
    // Example: output the name of the worksheet
    System.Console.WriteLine("Shape found in worksheet: " + targetWorksheet.Name);
}
else
{
    System.Console.WriteLine("Shape not found in any worksheet.");
}

// (Optional) Save the workbook if any modifications were made
// workbook.Save("output.xlsx");
