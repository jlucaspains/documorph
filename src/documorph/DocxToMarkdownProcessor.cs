using DocumentFormat.OpenXml.Linq;

namespace documorph;

/// <summary>
/// Converts a .docx file to markdown. The file is opened in read-only mode with file share capabilities.
/// </summary>
/// <param name="inputFile">The .docx file to be converted</param>
/// <exception cref="FileNotFoundException">Thrown when the input file does not exist</exception>
/// <exception cref="InvalidDataException">Thrown when the input file is not a valid .docx file</exception>
public sealed class DocxToMarkdownProcessor(string inputFile)
{
    public (string Result, IEnumerable<MediaModel> Media) Process()
    {
        using var wordPackage = Package.Open(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var wordDocument = WordprocessingDocument.Open(wordPackage);

        var model = DocumentModel.FromDocument(wordDocument);
        var result = new StringBuilder();

        var children = model.Children.Select((x, i) => (Element: x, Index: i)).ToList();
        foreach (var (element, index) in children)
        {
            switch (element)
            {
                case ParagraphModel paragraphModel:
                    var nextParagraphModel = index < children.Count - 1 ? children[index + 1].Element as ParagraphModel : null;
                    ProcessParagraph(paragraphModel, nextParagraphModel, false, result);
                    break;
                case TableModel tableModel:
                    ProcessTable(tableModel, result);
                    break;
            }
        }

        return (result.ToString(), model.Media);
    }

    private void ProcessParagraph(ParagraphModel paragraphModel, ParagraphModel? nextParagraphModel, bool skipNewLine, StringBuilder builder)
    {
        if (paragraphModel.IsHeading)
        {
            builder.Append($"{new string('#', paragraphModel.HeadingLevel)} ");
        }

        if (paragraphModel.IsList)
        {
            var listMarker = paragraphModel.ListFormat switch
            {
                "decimal" => "1.",
                "bullet" => "-",
                _ => ""
            };
            builder.Append($"{new string(' ', paragraphModel.ListItemLevel * 3)}{listMarker} ");
        }

        if (paragraphModel.IsQuote)
        {
            builder.Append($"> ");
        }

        var children = paragraphModel.Children.Select((x, i) => (Element: x, Index: i)).ToList();
        foreach (var (element, index) in children)
        {
            switch (element)
            {
                case RunModel run:
                    var previousRunModel = index > 0 ? children[index - 1].Element as RunModel : null;
                    var nextRunModel = index < children.Count - 1 ? children[index + 1].Element as RunModel : null;
                    ProcessRun(run, previousRunModel, nextRunModel, builder);
                    break;
                case HyperlinkModel hyperlink:
                    ProcessHyperlink(hyperlink, builder);
                    break;
            }
        }

        if (!skipNewLine && !paragraphModel.IsEmpty)
        {
            builder.AppendLine();
        }
        if (!skipNewLine && 
            (!paragraphModel.IsList || (!nextParagraphModel?.IsList ?? true)) &&
            !paragraphModel.IsEmpty)
        {
            builder.AppendLine();
        }
    }

    private void ProcessTable(TableModel tableModel, StringBuilder builder)
    {
        var isFirstRow = true;

        foreach (var row in tableModel.Rows)
        {
            foreach (var cell in row)
            {
                builder.Append('|');

                ProcessParagraph(cell, null, true, builder);
            }

            builder.AppendLine("|");

            if (isFirstRow)
            {
                builder.AppendLine("|" + string.Join("", Enumerable.Repeat("---|", row.Count())));
                isFirstRow = false;
            }
        }

        builder.AppendLine();
    }

    private void ProcessRun(RunModel runModel, RunModel? previousRunModel, RunModel? nextRunModel, StringBuilder builder)
    {
        if (runModel.IsBreak)
        {
            builder.AppendLine();
            return;
        }

        var markdownSymbols = new Dictionary<Func<RunModel, RunModel?, bool>, string>
        {
            { (rm, reference) => rm.IsBold && !(reference?.IsBold ?? false), "**" },
            { (rm, reference) => rm.IsItalic && !(reference?.IsItalic ?? false), "*" },
            { (rm, reference) => rm.IsStrike && !(reference?.IsStrike ?? false), "~~" },
            { (rm, reference) => rm.IsUnderline && !(reference?.IsUnderline ?? false), "__" }
        };

        if (runModel.IsImage)
        {
            builder.Append($"![");
        }

        var text = runModel.IsText ? runModel.Text : runModel.Image?.Description;

        foreach (var symbol in markdownSymbols)
        {
            if (symbol.Key(runModel, previousRunModel))
                builder.Append(symbol.Value);
        }

        builder.Append(text);

        foreach (var symbol in markdownSymbols)
        {
            if (symbol.Key(runModel, nextRunModel))
                builder.Append(symbol.Value);
        }

        if (runModel.IsImage)
        {
            builder.Append($"]({runModel.Image?.FileName})");
        }
    }

    private void ProcessHyperlink(HyperlinkModel hyperlink, StringBuilder builder)
    {
        builder.Append('[');

        var children = hyperlink.Runs.Select((x, i) => (Element: x, Index: i)).ToList();
        foreach (var (element, index) in children)
        {
            var previousRunModel = index > 0 ? children[index - 1].Element : null;
            var nextRunModel = index < children.Count - 1 ? children[index + 1].Element : null;
            ProcessRun(element, previousRunModel, nextRunModel, builder);
        }

        builder.Append($"]({hyperlink.Uri})");
    }
}
