namespace documorph;

public interface IDocumentChildren { }

public sealed class DocumentModel(IEnumerable<MediaModel> media, IEnumerable<IDocumentChildren> children)
{
    public IEnumerable<MediaModel> Media { get; private set; } = media;
    public IEnumerable<IDocumentChildren> Children { get; private set; } = children;
    
    public static DocumentModel FromDocument(WordprocessingDocument document)
    {
        var body = document?.MainDocumentPart?.Document.Body;
        var numberingDefinitionsPart = document?.MainDocumentPart?.NumberingDefinitionsPart;
        var hyperlinkRelationships = document?.MainDocumentPart?.HyperlinkRelationships
            ?? [];
        var parts = document?.MainDocumentPart?.Parts
            ?? [];
        var children = new List<IDocumentChildren>();

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

        foreach (var element in body?.ChildElements ?? new())
        {
            switch (element)
            {
                case Paragraph paragraph:
                    children.Add(ParagraphModel.FromParagraph(paragraph, numberingDefinitionsPart, hyperlinkRelationships, parts));
                    break;
                case Table table:
                    children.Add(TableModel.FromTable(table, hyperlinkRelationships));
                    break;
            }
        }

        var documentModel = new DocumentModel(media, children);

        return documentModel;
    }

    private static byte[] GetBytesFromStream(Stream stream)
    {
        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        return memoryStream.ToArray();
    }
}