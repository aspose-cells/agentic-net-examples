// Title: How to use foreach smart markers in Aspose.Cells C# to import a variable‑length list of product reviews
// Common Searches: Aspose.Cells foreach smart marker with List<Review> example in C# | populate Excel template rows from a POCO collection using WorkbookDesigner | C# generate product review report with Aspose.Cells smart markers | bind a variable‑length collection to a smart marker name and expand rows automatically | load template containing ${Reviews:ProductName} markers and export to Excel
// Tags: foreach smart markers with WorkbookDesigner | bind POCO collection to Aspose.Cells smart marker | dynamic row generation in Excel using Aspose.Cells | C# import variable length collection into Excel template | product review report generation Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace SmartMarkerDemo
{
    // Define a simple POCO to represent a product review
    // The example loads an Excel template that contains foreach smart markers, creates a List<Review> with product name, rating, and comment, binds this collection to the smart marker name "Reviews" via WorkbookDesigner, processes the markers to repeat the row for each review, and saves the resulting workbook as ReviewsReport.xlsx.
    public class Review
    {
        public string ProductName { get; set; }
        public int Rating { get; set; }          // e.g., 1‑5 stars
        public string Comment { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the Excel template that contains smart markers.
            // The template should have a row with foreach smart markers like:
            // ${Reviews:ProductName}   ${Reviews:Rating}   ${Reviews:Comment}
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Prepare a variable‑length collection of reviews.
            List<Review> reviews = new List<Review>
            {
                new Review { ProductName = "Widget A", Rating = 5, Comment = "Excellent!" },
                new Review { ProductName = "Gadget B", Rating = 4, Comment = "Very good, but could be cheaper." },
                new Review { ProductName = "Device C", Rating = 3, Comment = "Average performance." }
                // Add as many items as needed; the foreach smart marker will expand rows automatically.
            };

            // Bind the collection to the smart marker name "Reviews".
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Reviews", reviews);

            // Process the smart markers – the foreach syntax will repeat the row for each review.
            designer.Process();

            // Save the populated workbook.
            workbook.Save("ReviewsReport.xlsx");
        }
    }
}
