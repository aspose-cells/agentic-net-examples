# Working With Images Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Working With Images


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Working With Images**.

Example:

create-a-workbook.cs


## Required Namespaces

Most examples will require:

using Aspose.Cells;


## Common Pattern

Typical Aspose.Cells workflow:

Workbook workbook = new Workbook();

Worksheet sheet = workbook.Worksheets[0];

Cells cells = sheet.Cells;


## Output

Examples may generate:

- XLSX files
- PDF files
- CSV files
- Images

Output files are written to the working directory.
- convert-an-entire-workbook-to-a-multipage-tiff-using-default-rendering-options.cs
- render-a-workbook-as-tiff-using-300-dpi-resolution-to-improve-image-clarity.cs
- render-a-workbook-to-tiff-with-horizontal-and-vertical-resolutions-set-to-150-dpi.cs
- convert-a-workbook-to-tiff-using-eightbit-color-depth-for-smaller-output-files.cs
- convert-a-workbook-to-tiff-using-twentyfourbit-color-depth-for-highquality-images.cs
- generate-a-multipage-tiff-from-the-first-three-worksheets-of-a-loaded-workbook.cs
- track-workbooktotiff-conversion-progress-by-handling-the-conversion-progress-event-and-logging-percentages.cs
- render-a-workbook-to-tiff-and-write-the-result-into-a-memory-stream-for-further-processing.cs
- save-a-generated-tiff-file-directly-to-a-network-share-path-after-converting-the-workbook.cs
- embed-the-resulting-tiff-image-into-an-html-page-using-an-img-tag-with-appropriate-source-attribute.cs
- measure-conversion-duration-using-stopwatch-and-log-elapsed-time-for-performance-analysis.cs
