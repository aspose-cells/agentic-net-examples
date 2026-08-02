using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkersDemo
{
    // Sample data class representing a product review
    public class Review
    {
        public string Reviewer { get; set; }
        public int Rating { get; set; }
        public string Comment { get; set; }
        public DateTime Date { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a template workbook in memory and define smart markers
            // ------------------------------------------------------------
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Reviewer");
            cells["B1"].PutValue("Rating");
            cells["C1"].PutValue("Comment");
            cells["D1"].PutValue("Date");

            // Data row using foreach syntax (&=Collection.Column)
            // This row will be repeated for each item in the "Reviews" collection
            cells["A2"].PutValue("&=Reviews.Reviewer");
            cells["B2"].PutValue("&=Reviews.Rating");
            cells["C2"].PutValue("&=Reviews.Comment");
            cells["D2"].PutValue("&=Reviews.Date");

            // ------------------------------------------------------------
            // 2. Prepare a variable‑length collection of product reviews
            // ------------------------------------------------------------
            List<Review> reviews = new List<Review>
            {
                new Review
                {
                    Reviewer = "Alice",
                    Rating = 5,
                    Comment = "Excellent product!",
                    Date = DateTime.Today.AddDays(-2)
                },
                new Review
                {
                    Reviewer = "Bob",
                    Rating = 4,
                    Comment = "Very good, but could be cheaper.",
                    Date = DateTime.Today.AddDays(-1)
                },
                new Review
                {
                    Reviewer = "Charlie",
                    Rating = 3,
                    Comment = "Average experience.",
                    Date = DateTime.Today
                }
                // Add more reviews as needed – the foreach smart marker will handle any length
            };

            // ------------------------------------------------------------
            // 3. Initialize WorkbookDesigner with the template
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(template);

            // Bind the collection to the smart marker name "Reviews"
            designer.SetDataSource("Reviews", reviews);

            // Process the smart markers – the foreach row will be expanded automatically
            designer.Process();

            // ------------------------------------------------------------
            // 4. Save the populated workbook
            // ------------------------------------------------------------
            // Save to a file (you can also save to a stream if required)
            designer.Workbook.Save("ProductReviews.xlsx");
        }
    }
}