// Title: Generate GitHub‑Flavored Markdown Tables from an Aspose.Cells Workbook (C#)
// Description: This example creates a workbook, populates two worksheets with product and sales data, configures MarkdownSaveOptions to use the first row as the table header, align columns with spaces, and keep tables intact, then saves the result as a GitHub‑flavored Markdown file (DocumentationTables.md).
// Keywords: Aspose.Cells markdown export | C# generate markdown tables from Excel | GitHub flavored markdown Aspose.Cells | MarkdownSaveOptions table header | convert worksheet to markdown | Aspose.Cells documentation generation | Excel to markdown C# | space padded markdown columns
// Common Searches: How to export Aspose.Cells workbook to GitHub markdown tables in C# | Aspose.Cells MarkdownSaveOptions first row header example | Create markdown documentation from Excel worksheets using Aspose.Cells | Save Excel data as markdown without splitting tables | C# code to generate GitHub‑flavored markdown from multiple sheets
// Developer Intent: Export workbook worksheets as GitHub‑flavored markdown tables for documentation purposes.
// Use Cases: Generate product and sales tables for a project README automatically from Excel data. | Integrate markdown table generation into CI pipelines to keep technical docs synchronized with source data. | Produce consistently formatted markdown tables with header detection and column padding for developer guides.
// AI Prompts: Add a third worksheet and include its data as an additional markdown table in the same file. | Show how to set left, center, or right alignment for columns using MarkdownSaveOptions. | Provide code to insert a caption above each generated markdown table.

using System;
using Aspose.Cells;
using Aspose.Cells.Markdown;

namespace AsposeCellsExamples
{
    // This example creates a workbook, populates two worksheets with product and sales data, configures MarkdownSaveOptions to use the first row as the table header, align columns with spaces, and keep tables intact, then saves the result as a GitHub‑flavored Markdown file (DocumentationTables.md).
    public class GenerateMarkdownTables
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Populate first worksheet with sample product data
                // -------------------------------------------------
                Worksheet productsSheet = workbook.Worksheets[0];
                productsSheet.Name = "Products";

                // Header row
                productsSheet.Cells["A1"].PutValue("Product");
                productsSheet.Cells["B1"].PutValue("Price");
                productsSheet.Cells["C1"].PutValue("Quantity");

                // Data rows
                productsSheet.Cells["A2"].PutValue("Apple");
                productsSheet.Cells["B2"].PutValue(1.20);
                productsSheet.Cells["C2"].PutValue(50);

                productsSheet.Cells["A3"].PutValue("Banana");
                productsSheet.Cells["B3"].PutValue(0.80);
                productsSheet.Cells["C3"].PutValue(100);

                productsSheet.Cells["A4"].PutValue("Cherry");
                productsSheet.Cells["B4"].PutValue(2.50);
                productsSheet.Cells["C4"].PutValue(30);

                // -------------------------------------------------
                // Add a second worksheet with sales data
                // -------------------------------------------------
                Worksheet salesSheet = workbook.Worksheets.Add("Sales");

                // Header row
                salesSheet.Cells["A1"].PutValue("Region");
                salesSheet.Cells["B1"].PutValue("Sales");

                // Data rows
                salesSheet.Cells["A2"].PutValue("North");
                salesSheet.Cells["B2"].PutValue(15000);

                salesSheet.Cells["A3"].PutValue("South");
                salesSheet.Cells["B3"].PutValue(12000);

                salesSheet.Cells["A4"].PutValue("East");
                salesSheet.Cells["B4"].PutValue(13000);

                salesSheet.Cells["A5"].PutValue("West");
                salesSheet.Cells["B5"].PutValue(11000);

                // -------------------------------------------------
                // Configure Markdown save options
                // -------------------------------------------------
                MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
                {
                    // Use the first row of each sheet as the table header
                    TableHeaderType = MarkdownTableHeaderType.FirstRow,

                    // Align columns with spaces for better readability
                    AlignColumnPadding = ' ',

                    // Do not split tables by blank rows
                    SplitTablesByBlankRow = false
                };

                // -------------------------------------------------
                // Save the workbook as a GitHub‑flavored Markdown file
                // -------------------------------------------------
                workbook.Save("DocumentationTables.md", mdOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating markdown tables: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            GenerateMarkdownTables.Run();
        }
    }
}
