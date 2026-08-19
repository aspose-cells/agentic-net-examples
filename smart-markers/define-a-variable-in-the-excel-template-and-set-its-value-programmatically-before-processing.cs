// Title: Aspose.Cells C# – Define and Use Variables with Smart Markers in an Excel Template
// Description: Demonstrates how to create a "Variables" worksheet, assign name/value pairs (e.g., ReportTitle = Quarterly Sales), configure WorkbookDesigner to read that sheet, insert a smart marker "&=$ReportTitle" in a template sheet, process the markers, and save the populated workbook.
// Keywords: Aspose.Cells C# example | smart markers variables | WorkbookDesigner variables worksheet | set variable programmatically | Excel template variable replacement | Aspose.Cells variable sheet | C# Excel report title smart marker
// Common Searches: Aspose.Cells define variable in worksheet | C# smart markers variable replacement | WorkbookDesigner use variables sheet | set variable value before processing Aspose.Cells | how to use &=$ variable syntax in Aspose.Cells
// Developer Intent: Create a variables sheet, map name/value pairs, and have smart markers automatically substitute those values during processing.
// Use Cases: Insert a dynamic report title into the header of generated Excel reports. | Maintain a single source of constants (date, author, company) that can be referenced across multiple template sheets. | Load variable values from external sources (database, JSON, XML) into the variables worksheet before processing smart markers.
// AI Prompts: Generate C# code that adds multiple name/value pairs to a "Variables" worksheet and references each with smart markers on different template sheets using Aspose.Cells. | Explain how to read a JSON file, populate the variables worksheet with its data, and then process smart markers in Aspose.Cells. | Show how to change the variables worksheet name at runtime and still have WorkbookDesigner replace the smart markers correctly.

using System;
using Aspose.Cells;

namespace AsposeCellsVariableDemo
{
    // Demonstrates how to create a "Variables" worksheet, assign name/value pairs (e.g., ReportTitle = Quarterly Sales), configure WorkbookDesigner to read that sheet, insert a smart marker "&=$ReportTitle" in a template sheet, process the markers, and save the populated workbook.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // 2. Add a worksheet that will hold the variable definitions
            Worksheet variablesSheet = workbook.Worksheets.Add("Variables");
            // Define a variable name in column A and its value in column B
            variablesSheet.Cells["A1"].PutValue("ReportTitle");   // variable name
            variablesSheet.Cells["B1"].PutValue("Quarterly Sales"); // variable value

            // 3. Create a WorkbookDesigner and assign the workbook (lifecycle: create)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // 4. Tell the designer which worksheet contains the variables
            designer.VariablesWorksheetName = "Variables";

            // 5. Add a template worksheet that uses the variable via a smart marker
            Worksheet templateSheet = workbook.Worksheets.Add("Template");
            // Smart marker that references the variable defined above
            templateSheet.Cells["A1"].PutValue("&=$ReportTitle");

            // 6. Process the smart markers – the variable value will replace the marker
            designer.Process();

            // 7. Save the resulting workbook (lifecycle: save)
            workbook.Save("VariableDemo_Output.xlsx");
        }
    }
}
