// Title: Toggle Smart‑Marker Section Visibility with a Boolean Variable Using WorkbookDesigner.SetVariable in Aspose.Cells for .NET
// Description: Demonstrates how to create an Excel workbook, embed a conditional smart‑marker block (&if/$ShowSection … &endif), assign a boolean flag via WorkbookDesigner.SetVariable, process the markers, and save the file. Ideal for generating reports with optional sections.
// Keywords: Aspose.Cells SetVariable C# | smart markers conditional visibility | if endif smart marker .NET | WorkbookDesigner SetVariable example | toggle Excel template sections | dynamic report generation C#
// Common Searches: Aspose.Cells hide smart marker block with boolean | WorkbookDesigner.SetVariable usage C# | conditional smart markers example Aspose.Cells | how to show or hide sections in Excel template using Aspose | C# smart marker &if condition
// Developer Intent: Apply a boolean variable to control the display of a smart‑marker block and generate the final workbook.
// Use Cases: Create financial statements where footnotes appear only when a flag is true. | Build marketing dashboards that include optional analysis sections based on user selection. | Produce contract documents that show or hide clauses depending on configuration settings.
// AI Prompts: Show a C# code snippet that uses WorkbookDesigner.SetVariable to control &if smart markers in Aspose.Cells. | Explain when to prefer SetVariable over SetDataSource for scalar values in Aspose.Cells. | Provide an example of toggling multiple smart‑marker sections with a single boolean variable in .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create an Excel workbook, embed a conditional smart‑marker block (&if/$ShowSection … &endif), assign a boolean flag via WorkbookDesigner.SetVariable, process the markers, and save the file. Ideal for generating reports with optional sections.
class SmartMarkerVisibilityDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Template";

            // Define a conditional smart marker block
            sheet.Cells["A1"].PutValue("&if($ShowSection)");
            sheet.Cells["A2"].PutValue("This content is visible when ShowSection = true");
            sheet.Cells["A3"].PutValue("&endif");

            // Associate the workbook with a designer
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the variable that controls visibility (use SetDataSource for scalar values)
            designer.SetDataSource("ShowSection", true);

            // Process smart markers (populate data and evaluate conditions)
            designer.Process();

            // Save the workbook
            string outputPath = "SmartMarkerVisibility.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
