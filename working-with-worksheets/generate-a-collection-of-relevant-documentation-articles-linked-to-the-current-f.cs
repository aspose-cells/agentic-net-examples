using System;
using System.Collections.Generic;

namespace AsposeCellsDocs
{
    class Program
    {
        static void Main()
        {
            // Create a collection of documentation articles related to the current feature
            var docs = new List<DocumentationArticle>
            {
                // Built‑in document property that indicates hyperlink status
                new DocumentationArticle(
                    "BuiltInDocumentPropertyCollection.LinksUpToDate Property",
                    "https://docs.aspose.com/cells/net/builtindocumentpropertycollection-linksuptodate/"),

                // General information about the BuiltInDocumentPropertyCollection class
                new DocumentationArticle(
                    "BuiltInDocumentPropertyCollection Class",
                    "https://docs.aspose.com/cells/net/builtindocumentpropertycollection/"),

                // Information about custom document properties (often used together)
                new DocumentationArticle(
                    "CustomDocumentPropertyCollection Class",
                    "https://docs.aspose.com/cells/net/customdocumentpropertycollection/"),

                // How to access built‑in properties from a Workbook instance
                new DocumentationArticle(
                    "Workbook.BuiltInDocumentProperties",
                    "https://docs.aspose.com/cells/net/workbook-builtindocumentproperties/"),

                // Working with hyperlinks in worksheets (relevant to LinksUpToDate)
                new DocumentationArticle(
                    "Worksheet Hyperlinks",
                    "https://docs.aspose.com/cells/net/worksheet-hyperlinks/"),

                // Linked custom properties – another way to keep data in sync
                new DocumentationArticle(
                    "DocumentProperty.IsLinkedToContent Property",
                    "https://docs.aspose.com/cells/net/documentproperty-islinkedtocontent/"),

                // Using Aspose.Cells AI to summarize a spreadsheet (related to document metadata)
                new DocumentationArticle(
                    "Aspose.Cells AI – SpreadsheetSummarize",
                    "https://docs.aspose.com/cells/net/spreadsheetsummarize/")
            };

            // Output the list for easy navigation
            Console.WriteLine("Relevant Aspose.Cells documentation articles:");
            foreach (var article in docs)
            {
                Console.WriteLine($"{article.Title} - {article.Url}");
            }
        }
    }

    // Simple data transfer object to hold article title and URL
    public class DocumentationArticle
    {
        public string Title { get; }
        public string Url { get; }

        public DocumentationArticle(string title, string url)
        {
            Title = title;
            Url = url;
        }
    }
}