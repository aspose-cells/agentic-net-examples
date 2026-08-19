// Title: C# foreach smart markers for variable‑length collections in Aspose.Cells
// Description: This example builds an in‑memory workbook template, adds smart markers for a product name and a foreach collection of review texts, binds a product object and a List<Review> to the markers with WorkbookDesigner, processes the markers with LineByLine disabled, and saves the result as ProductReviews.xlsx.
// Keywords: Aspose.Cells | C# smart markers | foreach smart marker | variable length collection | WorkbookDesigner SetDataSource list | export reviews to Excel | dynamic rows Aspose.Cells | template workbook smart markers
// Common Searches: Aspose.Cells foreach smart marker example | bind List<T> to smart markers C# | variable length collection in Excel using Aspose.Cells | how to use LineByLine false with smart markers | export product reviews to Excel Aspose
// Developer Intent: Create an Excel sheet that automatically expands to list all reviews for a given product using foreach smart markers.
// Use Cases: Generate a product feedback report with a dynamic number of customer comments. | Export order invoices where each order contains a different count of line items. | Produce survey results where respondents may have varying numbers of answers.
// AI Prompts: Add a review date column to the foreach smart marker and show the updated code. | Show a nested foreach smart marker that lists multiple products, each with its own reviews. | Explain the impact of setting LineByLine = true versus false when processing smart markers.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// This example builds an in‑memory workbook template, adds smart markers for a product name and a foreach collection of review texts, binds a product object and a List<Review> to the markers with WorkbookDesigner, processes the markers with LineByLine disabled, and saves the result as ProductReviews.xlsx.
class Program
{
    static void Main()
    {
        // Create a template workbook in memory
        Workbook template = new Workbook();
        Worksheet sheet = template.Worksheets[0];

        // Add header cells
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Review");

        // Smart marker for a single product name
        sheet.Cells["A2"].PutValue("&=Product.Name");

        // Smart marker using foreach syntax to iterate over a variable‑length collection of reviews
        sheet.Cells["B2"].PutValue("&=Reviews.REVIEW_TEXT");

        // Define the range that contains the smart markers (required when LineByLine = false)
        sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

        // Prepare data sources
        var product = new { Name = "SuperWidget" };
        var reviews = new List<Review>
        {
            new Review { REVIEW_TEXT = "Excellent!" },
            new Review { REVIEW_TEXT = "Good value." },
            new Review { REVIEW_TEXT = "Could be better." }
        };

        // Initialize WorkbookDesigner with the template workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = template;
        designer.LineByLine = false; // Use range smart markers instead of line‑by‑line processing

        // Bind data sources to the smart marker names
        designer.SetDataSource("Product", product);
        designer.SetDataSource("Reviews", reviews);

        // Process the smart markers and populate the worksheet
        designer.Process();

        // Save the populated workbook
        designer.Workbook.Save("ProductReviews.xlsx");
    }

    // Simple class representing a product review
    public class Review
    {
        public string REVIEW_TEXT { get; set; }
    }
}
