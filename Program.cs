using System;
using System.IO;
using Microsoft.Office.Interop.OneNote;
using System.Xml;

namespace OneNoteMDImporter
{
    internal class Program
    {
        private static readonly double DefaultX = 36.0;
        private static readonly double DefaultY = 72.0;
        private static readonly double DefaultWidth = 600.0;
        private static readonly double DefaultHeight = 400.0;

        static void Main(string[] args)
        {
            // コマンドライン引数: mdfilepath, notebook, section, x, y, width, height
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: OneNoteMDImporter.exe <mdfilepath> <notebook> <section> [x] [y] [width] [height]");
                return;
            }
            string mdfilepath = args[0];
            string notebook = args[1];
            string section = args[2];
            double x = args.Length > 3 ? double.Parse(args[3]) : DefaultX;
            double y = args.Length > 4 ? double.Parse(args[4]) : DefaultY;
            double width = args.Length > 5 ? double.Parse(args[5]) : DefaultWidth;
            double height = args.Length > 6 ? double.Parse(args[6]) : DefaultHeight;

            // mdfilepath フルパス変換
            mdfilepath = Path.GetFullPath(mdfilepath);
            string body = MarkdownOperator.ConvertMarkdownToHtml(mdfilepath, File.ReadAllText(mdfilepath));
            string title = Path.GetFileNameWithoutExtension(mdfilepath);
            string style = File.ReadAllText("style.css");

            Func<string, string> contentXmlBuilder = (id) => BuildPageContentXml(id, x, y, width, height, title, style, body);
            OneNoteOperator.CreatePageInSection(notebook, section, contentXmlBuilder);
        }

        // テンプレート生成関数
        // @ref https://github.com/stevencohn/OneMore/blob/main/OneMore/Commands/Edit/ConvertMarkdownCommand.cs
        private static string BuildPageContentXml(string id, double px, double py, double w, double h, string t, string s, string b)
        {
            return $@"
<one:Page xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote' ID='{id}'>
  <one:Title>
    <one:OE>
      <one:T>{t}</one:T>
    </one:OE>
  </one:Title>
  <one:Outline>
    <one:Position x='{px}' y='{py}'/>
    <one:Size width='{w}' height='{h}'/>
    <one:OEChildren>
        <one:HTMLBlock>
          <one:Data><![CDATA[
            <html>
              <head><style>{s}</style></head>
              <body>{b}</body>
            </html>
          ]]></one:Data>
        </one:HTMLBlock>
    </one:OEChildren>
  </one:Outline>
</one:Page>";
        }
    }
}
