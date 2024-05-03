using System.Text;
using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace documorph;

public sealed class RunModel(Run run, string hyperlinkUrl, IEnumerable<DocumentFormat.OpenXml.Packaging.IdPartPair> parts)
{
    private readonly Run _run = run;
    private readonly IEnumerable<DocumentFormat.OpenXml.Packaging.IdPartPair> _parts = parts;

    public void AppendMarkdown(StringBuilder builder)
    {
        var bold = false;
        var italic = false;
        var strike = false;
        var underline = false;

        foreach (var element in _run.ChildElements)
        {
            switch (element)
            {
                case RunProperties runProperties:
                    bold = runProperties.Bold != null;
                    strike = runProperties.Strike != null;
                    italic = runProperties.Italic != null;
                    underline = runProperties.Underline != null;

                    break;
                case Text text:
                    AppendText(builder, text.Text, bold, italic, strike, underline, hyperlinkUrl);
                    break;
                case Drawing drawing:
                    AppendImage(builder, drawing, _parts);
                    break;
            }
        }
    }

    private static void AppendImage(StringBuilder builder, Drawing drawing, IEnumerable<DocumentFormat.OpenXml.Packaging.IdPartPair> _parts)
    {
        var images = from graphic in drawing
                .Descendants<DocumentFormat.OpenXml.Drawing.Graphic>()
                     let graphicData = graphic.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().FirstOrDefault()
                     let pic = graphicData.ElementAt(0)
                     let nvdp = pic.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>().FirstOrDefault()
                     let blip = pic.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault()
                     join part in _parts on blip?.Embed?.Value equals part
                         .RelationshipId
                     let image = part.OpenXmlPart as DocumentFormat.OpenXml.Packaging.ImagePart
                     let fileName = image.Uri.ToString()[(image.Uri.ToString().LastIndexOf('/') + 1)..]
                     select new
                     {
                         Id = blip.Embed,
                         FileName = fileName,
                         mime = image.ContentType,
                         Name = nvdp?.Name?.Value,
                         Description = nvdp?.Description?.Value
                     };

        foreach (var image in images)
        {
            builder.Append($"![{image.Description.Replace("\n", " ")}]({image.FileName})");
        }
    }

    private static void AppendText(StringBuilder builder, string text, bool bold, bool italic, bool strike, bool underscore, string hyperlinkUrl)
    {
        var markdown = text;

        if (bold)
            markdown = $"**{markdown}**";

        if (italic)
            markdown = $"*{markdown}*";

        if (strike)
            markdown = $"~~{markdown}~~";

        if (underscore)
            markdown = $"__{markdown}__";

        if (!string.IsNullOrEmpty(hyperlinkUrl))
            markdown = $"[{markdown}]({hyperlinkUrl})";

        if (!string.IsNullOrEmpty(markdown))
            builder.Append(markdown);
    }
}
