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

        var body = wordDocument?.MainDocumentPart?.Document.Body;
        var numberingDefinitionsPart = wordDocument?.MainDocumentPart?.NumberingDefinitionsPart;
        var hyperlinkRelationships = wordDocument?.MainDocumentPart?.HyperlinkRelationships
            ?? [];
        var parts = wordDocument?.MainDocumentPart?.Parts
            ?? [];

        var media = (from graphic in body?
                .Descendants<DocumentFormat.OpenXml.Drawing.Graphic>()
                     let graphicData = graphic.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().FirstOrDefault()
                     let pic = graphicData.ElementAt(0)
                     let blip = pic.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault()
                     join part in parts on blip?.Embed?.Value equals part
                         .RelationshipId
                     let image = part.OpenXmlPart as ImagePart
                     let fileName = image.Uri.ToString()[(image.Uri.ToString().LastIndexOf('/') + 1)..]
                     select new MediaModel(blip?.Embed?.Value ?? string.Empty, fileName, GetBytesFromStream(image.GetStream()))).ToList();

        var lastParagraphWasList = false;
        var builder = new StringBuilder();

        foreach (var element in body?.ChildElements ?? new())
        {
            var skipNewLine = false;
            switch (element)
            {
                case Paragraph paragraph:
                    var paragraphModel = new ParagraphModel(paragraph, numberingDefinitionsPart, hyperlinkRelationships, parts);
                    paragraphModel.AppendMarkdown(builder, lastParagraphWasList);

                    lastParagraphWasList = paragraphModel.IsList;
                    skipNewLine = paragraphModel.IsEmpty;
                    break;
                case Table table:
                    if (lastParagraphWasList)
                        builder.AppendLine();

                    new TableModel(table, hyperlinkRelationships).AppendMarkdown(builder);
                    lastParagraphWasList = false;
                    break;
                default:
                    skipNewLine = true;
                    break;
            }

            if (!skipNewLine)
                builder.AppendLine();
        }

        return (builder.ToString(), media);
    }

    private static byte[] GetBytesFromStream(Stream stream)
    {
        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        return memoryStream.ToArray();
    }
}
