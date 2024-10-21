using DocumentFormat.OpenXml.Linq;

namespace documorph;

/// <summary>
/// Converts a .docx file to Wikimedia. The Docx file is opened in read-only mode with file share capabilities.
/// </summary>
/// <param name="inputFile">The .docx file to be converted</param>
/// <param name="mediaOutputRelativePath">The relative path to the media output directory</param>
/// <exception cref="FileNotFoundException">Thrown when the input file does not exist</exception>
/// <exception cref="InvalidDataException">Thrown when the input file is not a valid .docx file</exception>
public sealed class DocxToWikiProcessor(string inputFile, string mediaOutputRelativePath) : IDocxToTargetConverter
{
    /// <summary>
    /// Processes the DOCX file and converts it to the AsciiDoc format.
    /// </summary>
    /// <returns>
    /// A tuple containing the result as a string and an enumerable collection of <see cref="MediaModel"/> objects.
    /// </returns>
    public (string Result, IEnumerable<MediaModel> Media) Process()
    {
        using var wordPackage = Package.Open(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var wordDocument = WordprocessingDocument.Open(wordPackage);

        var model = DocumentModel.FromDocument(wordDocument, mediaOutputRelativePath);
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
            builder.Append($"{new string('=', paragraphModel.HeadingLevel)} ");
        }

        if (paragraphModel.IsList)
        {
            var listMarker = paragraphModel.ListFormat switch
            {
                "decimal" => new string('#', paragraphModel.ListItemLevel + 1),
                "bullet" => new string('*', paragraphModel.ListItemLevel + 1),
                _ => ""
            };
            builder.Append($"{listMarker} ");
        }

        if (paragraphModel.IsQuote)
        {
            builder.Append($"<blockquote>");
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

        if (paragraphModel.IsQuote)
        {
            builder.Append($"</blockquote>");
        }

        if (paragraphModel.IsHeading)
        {
            builder.Append($" {new string('=', paragraphModel.HeadingLevel)}");
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

        var columns = tableModel.Rows.First().Select(x => "1");
        builder.AppendLine("{|");

        foreach (var row in tableModel.Rows)
        {
            if (!isFirstRow)
                builder.AppendLine("|-");

            foreach (var cell in row)
            {
                builder.Append('|');

                ProcessParagraph(cell, null, true, builder);
                builder.AppendLine();
            }

            isFirstRow = false;
        }
        // builder.AppendLine();
        builder.AppendLine("|}");
    }

    private void ProcessRun(RunModel runModel, RunModel? previousRunModel, RunModel? nextRunModel, StringBuilder builder)
    {
        if (runModel.IsBreak)
        {
            builder.AppendLine();
            return;
        }

        var markdownSymbols = new Dictionary<Func<RunModel, RunModel?, bool>, (string Start, string End)>
        {
            { (rm, reference) => rm.IsBold && !(reference?.IsBold ?? false), ("'''", "'''") },
            { (rm, reference) => rm.IsItalic && !(reference?.IsItalic ?? false), ("''", "''") },
            { (rm, reference) => rm.IsStrike && !(reference?.IsStrike ?? false), ("<s>", "</s>") },
            { (rm, reference) => rm.IsUnderline && !(reference?.IsUnderline ?? false), ("<u>", "</u>") }
        };

        if (runModel.IsImage)
        {
            builder.Append($"[[File:{runModel.Image?.FileName}|");
        }

        var text = runModel.IsText ? runModel.Text : runModel.Image?.Description;
        text ??= string.Empty;

        foreach (var symbol in markdownSymbols)
        {
            if (symbol.Key(runModel, previousRunModel))
                builder.Append(symbol.Value.Start);
        }

        // move spaces at the end of text to a new string
        var spaces = text.Reverse().TakeWhile(char.IsWhiteSpace).ToList();
        builder.Append(text.TrimEnd(' '));

        foreach (var symbol in markdownSymbols.Reverse())
        {
            if (symbol.Key(runModel, nextRunModel))
                builder.Append(symbol.Value.End);
        }

        builder.Append(string.Join(string.Empty, spaces));

        if (runModel.IsImage)
        {
            builder.Append("]]");
        }
    }

    private void ProcessHyperlink(HyperlinkModel hyperlink, StringBuilder builder)
    {
        builder.Append($"[{hyperlink.Uri} ");

        var children = hyperlink.Runs.Select((x, i) => (Element: x, Index: i)).ToList();
        foreach (var (element, index) in children)
        {
            var previousRunModel = index > 0 ? children[index - 1].Element : null;
            var nextRunModel = index < children.Count - 1 ? children[index + 1].Element : null;
            ProcessRun(element, previousRunModel, nextRunModel, builder);
        }

        builder.Append("]");
    }
}
