using Microsoft.Office.Interop.OneNote;
using System;
using System.Runtime.InteropServices;
using System.Xml;

namespace OneNoteMDImporter
{
    public static class OneNoteOperator
    {
        public static void CreatePageInSection(string notebookName, string sectionName, Func<string, string> contentXmlBuilder)
        {
            Application oneNoteApp = null;

            try
            {
                oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

                // ノートブックのIDを取得
                string xml;
                oneNoteApp.GetHierarchy(null, HierarchyScope.hsPages, out xml);

                var doc = new XmlDocument();
                doc.LoadXml(xml);

                // 指定名のノートブックのIDを取得
                var notebookNode = doc.SelectSingleNode($"//one:Notebook[@name='{notebookName}']", GetNamespaceManager(doc));
                if (notebookNode == null)
                {
                    Console.WriteLine($"ノートブック '{notebookName}' が見つかりません。");
                    return;
                }
                string notebookId = notebookNode.Attributes["ID"].Value;

                // 指定名のセクションのIDを取得
                var sectionNode = doc.SelectSingleNode($"//one:Section[@name='{sectionName}']", GetNamespaceManager(doc));
                if (sectionNode == null)
                {
                    Console.WriteLine($"セクション '{sectionName}' が見つかりません。");
                    return;
                }
                string sectionId = sectionNode.Attributes["ID"].Value;

                // 新しいページを作成
                string newPageId;
                oneNoteApp.CreateNewPage(sectionId, out newPageId, NewPageStyle.npsDefault);

                // コールバック関数でページXMLを生成
                string pageXml = contentXmlBuilder(newPageId);
                oneNoteApp.UpdatePageContent(pageXml, DateTime.MinValue, XMLSchema.xs2013);

                Console.WriteLine("ページを作成しました。");
            }
            finally
            {
                if (oneNoteApp != null)
                {
                    Marshal.ReleaseComObject(oneNoteApp);
                    oneNoteApp = null;
                }

                // ガベージコレクションを強制的に実行（必要に応じて）
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static XmlNamespaceManager GetNamespaceManager(XmlDocument doc)
        {
            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");
            return nsmgr;
        }
    }
}
