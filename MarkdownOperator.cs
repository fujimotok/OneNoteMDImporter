using System;
using System.IO;
using Markdig;
using Markdig.Renderers;

namespace OneNoteMDImporter
{
    // @ref https://github.com/stevencohn/OneMore/blob/main/OneMore/Commands/File/Markdown/OneMoreDig.cs
    internal class MarkdownOperator
    {
        public static string ConvertMarkdownToHtml(string path, string markdown)
        {
            using (var writer = new StringWriter())
            {
                var renderer = new HtmlRenderer(writer)
                {
                    BaseUrl = new Uri($"{Path.GetDirectoryName(path)}/")
                };

                var pipeline = new MarkdownPipelineBuilder()
                    .UseAdvancedExtensions()
                    .Build();

                pipeline.Setup(renderer);

                var doc = Markdown.Parse(markdown, pipeline);

                renderer.Render(doc);
                writer.Flush();

                return writer.ToString();
            }
        }
    }
}
