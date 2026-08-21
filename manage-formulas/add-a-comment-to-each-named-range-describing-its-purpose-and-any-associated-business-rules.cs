// Title: Add descriptive comments to named ranges in Aspose.Cells using C#
// Description: Creates a workbook, defines two named ranges (ProductList and PriceList) on a "Data" sheet, attaches purpose‑and‑rule comments via the Name.Comment property, sorts the name collection for faster lookup, and saves the file as NamedRangesWithComments.xlsx.
// Keywords: Aspose.Cells named range comment C# | set Name.Comment Aspose.Cells | define named range with description .NET | sort defined names Aspose.Cells | save workbook with named range metadata
// Common Searches: how to add a comment to a named range Aspose.Cells C# | Aspose.Cells set business rule on defined name | sorting named ranges in a workbook Aspose.Cells | retrieve and modify Name.Comment property Aspose.Cells
// Developer Intent: Attach purpose and validation rules to each defined name in a workbook.
// Use Cases: Document the ProductList range for sales reporting and enforce a non‑empty rule. | Annotate the PriceList range with a rule that all prices must be positive. | Improve performance by sorting a large collection of named ranges before saving.
// AI Prompts: Generate C# code that loops through all named ranges in a workbook and assigns comments from a dictionary of descriptions using Aspose.Cells. | Show how to read, update, and persist the Comment of a specific named range in Aspose.Cells for .NET. | Explain a method to validate that every named range comment follows the pattern "Purpose: …; Rule: …" before calling workbook.Save.

using System;
using Aspose.Cells;

// Creates a workbook, defines two named ranges (ProductList and PriceList) on a "Data" sheet, attaches purpose‑and‑rule comments via the Name.Comment property, sorts the name collection for faster lookup, and saves the file as NamedRangesWithComments.xlsx.
class AddCommentsToNamedRanges
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it a meaningful name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate some sample data that will be referenced by named ranges
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["B2"].PutValue(1.2);
        sheet.Cells["B3"].PutValue(0.8);

        // Get the collection of defined names in the workbook
        NameCollection names = workbook.Worksheets.Names;

        // Define a named range for the list of products
        int prodIndex = names.Add("ProductList");
        Name productRange = names[prodIndex];
        productRange.RefersTo = "=Data!$A$2:$A$3";
        // Add a comment describing purpose and business rule
        productRange.Comment = "List of products used in sales reports. Must not be empty.";

        // Define a named range for the corresponding prices
        int priceIndex = names.Add("PriceList");
        Name priceRange = names[priceIndex];
        priceRange.RefersTo = "=Data!$B$2:$B$3";
        // Add a comment describing purpose and business rule
        priceRange.Comment = "Corresponding prices. Values must be positive numbers.";

        // Sort names for better performance when many names exist
        workbook.Worksheets.SortNames();

        // Save the workbook with the named ranges and their comments
        workbook.Save("NamedRangesWithComments.xlsx");
    }
}
