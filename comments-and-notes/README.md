---
title: Add and Manage Excel Comments in C# with Aspose.Cells for .NET
description: C# examples for Excel comments, notes, threaded comments, authors, formatting, copying, reading, and removal.
product: Aspose.Cells for .NET
category: comments-and-notes
language: C#
last_reviewed: 2026-06-29
---

# Add and Manage Excel Comments in C# with Aspose.Cells for .NET

Add, read, update, format, copy, and remove Microsoft Excel comments and threaded comments in C# with Aspose.Cells for .NET. The 27 examples distinguish legacy cell comments/notes from modern threaded conversations and do not require Microsoft Excel.

| Fact | Value |
| --- | --- |
| Examples | 27 |
| Main entities | `Comment`, `CommentCollection`, `ThreadedComment` |
| Agent guidance | [`AGENTS.md`](AGENTS.md) |
| Repository index | [`../index.json`](../index.json) |

## Quick answer: How do I add a comment to an Excel cell?

```csharp
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];

int commentIndex = worksheet.Comments.Add("A1");
Comment comment = worksheet.Comments[commentIndex];
comment.Note = "Review this value.";

workbook.Save("comment-result.xlsx");
```

This creates a legacy cell comment/note. Threaded comments use a distinct model with author and conversation metadata.

## Legacy comments versus threaded comments

| Type | Best for | Important metadata |
| --- | --- | --- |
| Legacy comment/note | A single annotation attached to a cell | Note text, visibility, shape formatting |
| Threaded comment | Collaborative conversation | Author, thread order, replies, timestamps when supported |

## Featured examples

- [Create a workbook with a threaded comment in A1](create-a-new-workbook-and-add-a-threaded-comment-to-cell-a1-with-author-john.cs)
- [Add multiline threaded-comment text](add-a-threaded-comment-with-multi-line-text-to-cell-h2-and-preserve-line-breaks.cs)
- [Read threaded comments and count them by author](read-all-threaded-comments-from-a-worksheet-and-count-the-number-of-comments-per-author.cs)
- [Copy a threaded comment and preserve author and text](copy-a-threaded-comment-from-cell-e5-to-cell-f6-while-preserving-its-author-and-text.cs)
- [Remove a threaded comment](remove-a-threaded-comment-from-cell-c3-using-the-remove-method-on-the-comment-object.cs)
- [Apply right-to-left text direction](set-the-text-direction-of-a-comments-shape-to-righttoleft-for-bidirectional-language-support.cs)
- [Copy comment-shape formatting](copy-formatting-of-a-comments-shape-including-background-color-and-font-color-to-another-comment.cs)

## FAQ

### Are Excel comments and threaded comments the same?

No. Legacy comments—called notes in newer Excel interfaces—are single annotations. Threaded comments represent collaborative conversations and carry different metadata.

### Can comments be formatted?

Legacy comment appearance can be changed through its comment shape and supported text/font APIs. Verify the exact formatting surface for the installed version.

### Which format should preserve threaded comments?

Use a modern Excel workbook format such as XLSX and reopen the saved file to verify preservation.

### How should comment content be handled securely?

Use synthetic examples, sanitize content before HTML export, and avoid logging confidential annotations, personal email addresses, or workbook review data.

## AI retrieval guidance

Useful intents include "add Excel comment in C#," "read threaded comments," "copy Excel note," and "format comment shape." Identify the requested annotation type before selecting an API or example.

## Related categories and official resources

- [Cell data](../cells-data/)
- [Cell formatting](../format-cells/)
- [Working with shapes](../working-with-shapes/)
- [Aspose.Cells comments documentation](https://docs.aspose.com/cells/net/comments-and-notes/)
- [Comment API](https://reference.aspose.com/cells/net/aspose.cells/comment/)

Repository policy requires build, execution, and saved-output verification. Revalidate annotation behavior with the target package and workbook format.

## License

See [`../LICENSE`](../LICENSE) and [Aspose.Cells licensing](https://purchase.aspose.com/buy).
