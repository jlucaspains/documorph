using DocumentFormat.OpenXml.Linq;

namespace documorph;

public sealed record ImageModel(string FileName, string MimeType, string? Description);

public sealed class RunModel(bool IsBold, bool isItalic, bool isStrike,
    bool isUnderline, bool isText, string? text, bool isImage,
    ImageModel? image, bool isBreak) : IParagraphChildModel
{
    public bool IsBold { get; private set; } = IsBold;
    public bool IsItalic { get; private set; } = isItalic;
    public bool IsStrike { get; private set; } = isStrike;
    public bool IsUnderline { get; private set; } = isUnderline;
    public bool IsText { get; private set; } = isText;
    public string? Text { get; private set; } = text;
    public bool IsImage { get; private set; } = isImage;
    public ImageModel? Image { get; private set; } = image;
    public bool IsBreak { get; private set; } = isBreak;

    public static RunModel FromRun(Run run, IEnumerable<IdPartPair> parts)
    {
        var isBold = false;
        var isItalic = false;
        var isStrike = false;
        var isUnderline = false;
        var isText = false;
        var text = string.Empty;
        var isImage = false;
        ImageModel? image = null;
        var isBreak = false;
        foreach (var element in run.ChildElements)
        {
            switch (element)
            {
                case RunProperties runProperties:
                    isBold = runProperties.Bold != null;
                    isItalic = runProperties.Italic != null;
                    isStrike = runProperties.Strike != null;
                    isUnderline = runProperties.Underline != null;

                    break;
                case Text textElement:
                    isText = true;
                    text = textElement.Text;
                    break;
                case Drawing drawing:
                    image = GetImageModel(drawing, parts);
                    isImage = image != null;
                    break;
                case Break:
                    isBreak = true;
                    break;
            }
        }

        return new RunModel(isBold, isItalic, isStrike, isUnderline, isText, text, isImage, image, isBreak);
    }

    private static ImageModel? GetImageModel(Drawing drawing, IEnumerable<IdPartPair> parts)
    {
        var images = from graphic in drawing
                .Descendants<DocumentFormat.OpenXml.Drawing.Graphic>()
                     let graphicData = graphic.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().FirstOrDefault()
                     let pic = (graphicData?.Any() ?? false) ? graphicData?.ElementAt(0) : null
                     let nvdp = pic?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>().FirstOrDefault()
                     let blip = pic?.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault()
                     join part in parts on blip?.Embed?.Value equals part
                         .RelationshipId
                     let image = part.OpenXmlPart as ImagePart
                     let fileName = image?.Uri.ToString()[(image.Uri.ToString().LastIndexOf('/') + 1)..]
                     let description = (nvdp?.Description?.Value ?? string.Empty).Replace("\n", " ")
                     select new ImageModel(fileName, image.ContentType, description);

        return images.FirstOrDefault();
    }
}
