// Title: Set the Author built‑in property of an Excel workbook using Aspose.Cells for .NET (C#)
// Description: Loads an existing .xlsx file with Aspose.Cells, assigns a contributor identifier to the workbook's BuiltInDocumentProperties.Author field, and saves the updated workbook.
// Keywords: Aspose.Cells C# set author | Excel built‑in document properties | Workbook Author property | update Excel metadata Aspose | C# Aspose.Cells document properties | set contributor ID author Excel
// Common Searches: how to set author property in Excel with Aspose.Cells | C# update built‑in document properties Aspose | change Excel file author using Aspose.Cells | save workbook after modifying metadata Aspose | Aspose.Cells set workbook author programmatically
// Developer Intent: Assign a specific contributor ID to the Author built‑in property of an existing Excel workbook and persist the change.
// Use Cases: Replace a placeholder author with the actual contributor ID before distributing a report. | Batch‑process generated workbooks to ensure the author field matches the responsible user. | Integrate author metadata updates into a document‑management workflow for compliance.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, sets the Author built‑in property from a variable, and saves the file. | Show how to update multiple built‑in properties (Author, Title, Subject) in a workbook using Aspose.Cells for .NET. | Provide error‑handling examples for missing input files and permission issues when setting the Author property with Aspose.Cells.

using System;
using Aspose.Cells;

namespace UpdateAuthorProperty
{
    // Loads an existing .xlsx file with Aspose.Cells, assigns a contributor identifier to the workbook's BuiltInDocumentProperties.Author field, and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that needs to be updated
            string inputFilePath = "input.xlsx";

            // Path where the updated workbook will be saved
            string outputFilePath = "output.xlsx";

            // Contributor identifier that will be set as the Author property
            string contributorId = "Contributor123";

            // Load the existing workbook from the file system
            Workbook workbook = new Workbook(inputFilePath);

            // Update the built‑in Author property with the contributor identifier
            workbook.BuiltInDocumentProperties.Author = contributorId;

            // Save the workbook with the updated property
            workbook.Save(outputFilePath);
        }
    }
}
