// Title: Replace placeholder in a named TextBox of an Excel workbook using Aspose.Cells for .NET (C#)
// Description: C# sample that loads an Excel workbook with Aspose.Cells, accesses the first worksheet, finds the TextBox called "HeaderBox", swaps the <TAG_1> token for a runtime title, and writes the modified file.
// Keywords: Aspose.Cells | C# | Excel TextBox | named TextBox | replace placeholder | dynamic title | shape text update | Aspose.Cells TextBox replace | Excel shape manipulation | programmatic Excel editing
// Common Searches: Aspose.Cells replace text in TextBox | C# change text of named TextBox in Excel | replace <TAG_1> in Excel shape using Aspose | update HeaderBox TextBox content Aspose.Cells | programmatically edit Excel TextBox text .NET | replace tag in Excel template Aspose.Cells
// Developer Intent: Swap the <TAG_1> token inside the HeaderBox TextBox with a dynamic title and save the workbook.
// Use Cases: Populate a quarterly‑report template by inserting the report title into a pre‑designed HeaderBox shape. | Automate branding of Excel templates by replacing placeholder tags in named TextBoxes with client‑specific values during generation. | Run a batch job that opens multiple workbooks, updates TextBox placeholders with runtime data, and saves each file.
// AI Prompts: Generate C# code with Aspose.Cells that locates a TextBox named 'HeaderBox' on the first worksheet and replaces the placeholder '<TAG_1>' with a variable string. | Create a reusable method taking workbook path, TextBox name, placeholder tag, and replacement value, then updates the TextBox text and saves the file using Aspose.Cells. | Explain how to handle missing TextBox or absent placeholder safely when performing text replacement in Excel shapes with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextBoxReplace
{
    // C# sample that loads an Excel workbook with Aspose.Cells, accesses the first worksheet, finds the TextBox called "HeaderBox", swaps the <TAG_1> token for a runtime title, and writes the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the TextBox named "HeaderBox"
            TextBox headerBox = worksheet.TextBoxes["HeaderBox"];

            if (headerBox != null && !string.IsNullOrEmpty(headerBox.Text))
            {
                // Dynamic title to replace the placeholder
                string dynamicTitle = "Quarterly Report 2026";

                // Replace the placeholder <TAG_1> with the dynamic title
                headerBox.Text = headerBox.Text.Replace("<TAG_1>", dynamicTitle);
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
