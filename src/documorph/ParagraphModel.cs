using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace documorph;

public sealed class ParagraphModel(Paragraph paragraph, NumberingDefinitionsPart? numberingDefinitionsPart, IEnumerable<HyperlinkRelationship> hyperlinkRelationships, IEnumerable<IdPartPair> parts)
{
    private readonly Paragraph _paragraph = paragraph;
    private readonly NumberingDefinitionsPart? _numberingDefinitionsPart = numberingDefinitionsPart;
    private readonly IEnumerable<HyperlinkRelationship> _hyperlinkRelationships = hyperlinkRelationships;
    private readonly IEnumerable<IdPartPair> _parts = parts;
    public bool IsList { get; private set; }
    public bool IsEmpty { get; private set; }

    public void AppendMarkdown(StringBuilder builder, bool lastParagraphWasList, bool isWithinAnotherElement = false)
    {
        var paragraphStyle = _paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var initialBuilderLength = builder.Length;

        var isHeading = false;
        var isList = false;
        var isQuote = false;

        if (paragraphStyle != null)
        {
            isHeading = paragraphStyle.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
            isList = paragraphStyle.StartsWith("ListParagraph", StringComparison.OrdinalIgnoreCase) && _paragraph.ParagraphProperties != null;
            isQuote = paragraphStyle.StartsWith("Quote", StringComparison.OrdinalIgnoreCase) && _paragraph.ParagraphProperties != null;
        }

        if (lastParagraphWasList && !isList)
            builder.AppendLine();

        if (isHeading && paragraphStyle != null)
            builder.Append($"{GetHeaderMarkdownPrefix(paragraphStyle)} ");

        if (isList)
            builder.Append($"{GetListMarkdownPrefix(_paragraph.ParagraphProperties!, _numberingDefinitionsPart)} ");

        if (isQuote)
            builder.Append("> ");

        IsList = isList;

        foreach (var element in _paragraph.ChildElements)
        {
            switch (element)
            {
                case Run run:
                    new RunModel(run, _parts).AppendMarkdown(builder);
                    break;
                case Hyperlink hyperlink:
                    new HyperlinkModel(hyperlink, _hyperlinkRelationships).AppendMarkdown(builder);
                    break;
            }
        }

        IsEmpty = builder.Length == initialBuilderLength;
        if (!isWithinAnotherElement && !IsList && !IsEmpty)
            builder.AppendLine();
    }

    private static string GetHeaderMarkdownPrefix(string headerName)
    {
        var headerLevelString = headerName.Replace("Heading", string.Empty, StringComparison.OrdinalIgnoreCase);
        _ = int.TryParse(headerLevelString, out var result);

        return new string('#', result);
    }

    private static string GetListMarkdownPrefix(ParagraphProperties paragraphProperties, NumberingDefinitionsPart? numberingDefinitionsPart)
    {
        var level = paragraphProperties.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
        var numberId = paragraphProperties.NumberingProperties?.NumberingId?.Val?.Value;
        var listAbstractNumberId = numberingDefinitionsPart?.Numbering?.NumberingDefinitionsPart?
            .Numbering?.Descendants<NumberingInstance>().FirstOrDefault(x => x.NumberID?.Value == numberId)?
            .Descendants<AbstractNumId>().FirstOrDefault()?.Val?.Value;

        var levelString = new string(' ', level * 3);

        var format = numberingDefinitionsPart?.Numbering?.Descendants<AbstractNum>()?.FirstOrDefault(x => x.AbstractNumberId?.Value == listAbstractNumberId)?.Descendants<Level>()?.FirstOrDefault()?.NumberingFormat?.Val?.InnerText;

        return format switch
        {
            "bullet" => $"{levelString}-",
            "decimal" => $"{levelString}1.",
            _ => string.Empty
        };
    }
}