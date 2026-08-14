// Title: C# – Batch Replace Placeholder Text in SmartArt Across All Worksheets with Aspose.Cells
// Description: Loads an Excel workbook, iterates every worksheet, finds SmartArt shapes, accesses their grouped shapes via GetResultOfSmartArt, replaces a specified placeholder (e.g., &=$CompanyName) with a new value, and saves the file using OoxmlSaveOptions.UpdateSmartArt to refresh cached graphics.
// Keywords: Aspose.Cells | C# SmartArt text replace | batch SmartArt update | replace placeholder in Excel SmartArt | GetResultOfSmartArt | GroupShape | UpdateSmartArt | multiple worksheets | Excel automation | Aspose.Cells .NET
// Common Searches: Aspose.Cells replace text in SmartArt | C# batch replace placeholder in Excel SmartArt | How to update SmartArt text in all sheets using Aspose.Cells | Replace &=$CompanyName in SmartArt with Aspose.Cells .NET | Iterate SmartArt shapes across workbook Aspose.Cells | Refresh SmartArt after text change Aspose.Cells | SmartArt text replacement example C#
// Developer Intent: Replace a placeholder string in every SmartArt shape throughout a workbook.
// Use Cases: Insert company name into SmartArt diagrams of a financial report template. | Localize SmartArt captions for multilingual Excel dashboards. | Apply brand‑wide text changes after a corporate rebrand. | Automate dynamic data insertion into SmartArt for monthly reporting. | Generate customized SmartArt for client‑specific presentations.
// AI Prompts: Write C# code using Aspose.Cells to search all worksheets for SmartArt shapes and replace a given token with a new value, ensuring UpdateSmartArt is enabled. | Explain the role of GetResultOfSmartArt and GroupShape when modifying SmartArt text in Aspose.Cells. | Create a robust batch routine that logs each replacement, handles missing placeholders, and catches exceptions during SmartArt processing. | Provide a PowerShell script that calls the compiled C# program to process multiple Excel files in a folder. | Suggest unit tests for verifying SmartArt text replacement logic with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtBatchReplace
{
    // Loads an Excel workbook, iterates every worksheet, finds SmartArt shapes, accesses their grouped shapes via GetResultOfSmartArt, replaces a specified placeholder (e.g., &=$CompanyName) with a new value, and saves the file using OoxmlSaveOptions.UpdateSmartArt to refresh cached graphics.
    class Program
    {
        static void Main()
        {
            const string templatePath = "TemplateWithSmartArt.xlsx";
            const string resultPath = "ResultWithSmartArtReplaced.xlsx";

            try
            {
                // Verify that the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the workbook that contains SmartArt objects
                Workbook workbook = new Workbook(templatePath);

                // Define the placeholder text to search for and its replacement
                string placeholder = "&=$CompanyName";   // example placeholder used in SmartArt
                string newValue = "Contoso Ltd.";

                // Perform batch replacement on SmartArt shapes across all worksheets
                ReplaceSmartArtText(workbook, placeholder, newValue);

                // Save the workbook with SmartArt update enabled
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true   // ensure cached SmartArt shapes are refreshed
                };
                workbook.Save(resultPath, saveOptions);

                Console.WriteLine($"SmartArt text replacement completed. Result saved to: {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="workbook">The workbook to process.</param>
        /// <param name="placeholder">The text to find inside SmartArt.</param>
        /// <param name="newValue">The text to replace the placeholder with.</param>
        static void ReplaceSmartArtText(Workbook workbook, string placeholder, string newValue)
        {
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Get the grouped shape representation of the SmartArt
                        GroupShape smartArtGroup = shape.GetResultOfSmartArt();

                        // Iterate through all grouped shapes that make up the SmartArt
                        foreach (Shape innerShape in smartArtGroup.GetGroupedShapes())
                        {
                            // If the shape contains text and matches the placeholder, replace it
                            if (!string.IsNullOrEmpty(innerShape.Text) && innerShape.Text.Contains(placeholder))
                            {
                                innerShape.Text = innerShape.Text.Replace(placeholder, newValue);
                            }
                        }
                    }
                }
            }
        }
    }
}
