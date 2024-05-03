using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace documorph;

public sealed class HyperlinkModel(Hyperlink hyperlink, IEnumerable<HyperlinkRelationship> hyperlinkRelationships)
{
    public void AppendMarkdown(StringBuilder builder)
    {
        var hyperlinkId = hyperlink.Id?.Value;
        var hyperlinkUrl = hyperlinkRelationships
            .FirstOrDefault(item => item.Id == hyperlinkId)?.Uri.ToString()
            ?? string.Empty;

        foreach (var element in hyperlink.ChildElements)
        {
            switch (element)
            {
                case Run run:
                    new RunModel(run, hyperlinkUrl, []).AppendMarkdown(builder);
                    break;
            }
        }
    }
}
