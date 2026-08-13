// Title: Replace <TAG_1> in a named TextBox (HeaderBox) with a dynamic title using Aspose.Cells for .NET
// Description: Loads an Excel workbook, accesses the first worksheet, retrieves the TextBox called "HeaderBox", substitutes the <TAG_1> placeholder with a runtime title, and saves the updated file.
// Keywords: Aspose.Cells | C# | Excel TextBox | named shape | replace placeholder | dynamic title | HeaderBox | load workbook | save workbook | worksheet shape text
// Common Searches: How to change the text of a TextBox named HeaderBox in Excel with Aspose.Cells | Aspose.Cells replace placeholder tag in Excel shape | C# replace <TAG_1> in an Excel TextBox | Update header title in an Excel template using Aspose.Cells | Get TextBox by name in Aspose.Cells .NET
// Developer Intent: Swap the <TAG_1> tag inside the HeaderBox TextBox for a generated title and write the changes back to the workbook.
// Use Cases: Automate quarterly report generation by inserting the report period into a pre‑designed header TextBox. | Populate invoice templates with customer‑specific data by replacing placeholder tags in shape text. | Refresh dashboard workbooks programmatically, updating the title TextBox across multiple files.
// AI Prompts: Write C# code with Aspose.Cells that locates a TextBox named "HeaderBox" and replaces a <TAG_1> placeholder with a variable string. | Provide an example that iterates over all TextBoxes in a worksheet and substitutes any <TAG_*> placeholders using values from a dictionary. | Explain how to safely handle cases where the specified TextBox does not exist or its Text property is null when using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, accesses the first worksheet, retrieves the TextBox called "HeaderBox", substitutes the <TAG_1> placeholder with a runtime title, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the TextBox named "HeaderBox" from the worksheet
        TextBox headerBox = worksheet.TextBoxes["HeaderBox"];

        // Ensure the TextBox exists and contains text
        if (headerBox != null && headerBox.Text != null)
        {
            // Define the dynamic title that will replace the placeholder
            string dynamicTitle = "Quarterly Report 2026";

            // Replace the placeholder <TAG_1> with the dynamic title
            headerBox.Text = headerBox.Text.Replace("<TAG_1>", dynamicTitle);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
