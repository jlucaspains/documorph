using DocumentFormat.OpenXml;

namespace documorph;

public interface IParagraphChildModel { }

public sealed class ParagraphModel(bool isList, string? listFormat, int listItemLevel,
    bool isHeading, int headingLevel, bool isQuote,
    IEnumerable<IParagraphChildModel> children): IDocumentChildren
{
    public bool IsList { get; private set; } = isList;
    public string? ListFormat { get; private set; } = listFormat;
    public int ListItemLevel { get; private set; } = listItemLevel;
    public bool IsHeading { get; private set; } = isHeading;
    public int HeadingLevel { get; private set; } = headingLevel;
    public bool IsQuote { get; private set; } = isQuote;
    public bool IsEmpty {get; private set; } = !children.Any();
    public IEnumerable<IParagraphChildModel> Children { get; private set; } = children;

    public static ParagraphModel FromParagraph(Paragraph paragraph, NumberingDefinitionsPart? numberingDefinitionsPart, IEnumerable<HyperlinkRelationship> hyperlinkRelationships, IEnumerable<MediaModel> parts)
    {
        var children = paragraph.ChildElements
            .Select<OpenXmlElement, IParagraphChildModel?>(element => element switch
            {
                Run run => RunModel.FromRun(run, parts),
                Hyperlink hyperlink => HyperlinkModel.FromHyperlink(hyperlink, hyperlinkRelationships),
                _ => null
            })
            .Where(child => child != null)
            .ToList();

        var paragraphStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

        var isHeading = false;
        var headingLevel = 0;
        var isList = false;
        var listFormat = string.Empty;
        var listItemLevel = 0;
        var isQuote = false;

        if (paragraphStyle != null)
        {
            isHeading = paragraphStyle.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
            isList = paragraphStyle.StartsWith("ListParagraph", StringComparison.OrdinalIgnoreCase) && paragraph.ParagraphProperties != null;
            isQuote = paragraphStyle.StartsWith("Quote", StringComparison.OrdinalIgnoreCase) && paragraph.ParagraphProperties != null;
        }
        
        if (isHeading && paragraphStyle != null)
            headingLevel = GetHeaderLevel(paragraphStyle);

        if (isList)
            (listFormat, listItemLevel) = GetListItemDetail(paragraph.ParagraphProperties!, numberingDefinitionsPart);



        return new ParagraphModel(isList, listFormat, listItemLevel, isHeading, headingLevel, isQuote, children!);
    }
    
    private static int GetHeaderLevel(string headerName)
    {
        var headerLevelString = headerName.Replace("Heading", string.Empty, StringComparison.OrdinalIgnoreCase);
        _ = int.TryParse(headerLevelString, out var result);

        return result;
    }

    private static (string Style, int Level) GetListItemDetail(ParagraphProperties paragraphProperties, NumberingDefinitionsPart? numberingDefinitionsPart)
    {
        var level = paragraphProperties.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
        var numberId = paragraphProperties.NumberingProperties?.NumberingId?.Val?.Value;
        var listAbstractNumberId = numberingDefinitionsPart?.Numbering?.NumberingDefinitionsPart?
            .Numbering?.Descendants<NumberingInstance>().FirstOrDefault(x => x.NumberID?.Value == numberId)?
            .Descendants<AbstractNumId>().FirstOrDefault()?.Val?.Value;

        var format = numberingDefinitionsPart?.Numbering?.Descendants<AbstractNum>()?
            .FirstOrDefault(x => x.AbstractNumberId?.Value == listAbstractNumberId)?
            .Descendants<Level>()?
            .FirstOrDefault()?
            .NumberingFormat?
            .Val?
            .InnerText;

        return (format ?? string.Empty, level);
    }
}