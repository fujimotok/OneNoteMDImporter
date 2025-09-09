using System;
using Microsoft.Office.Interop.OneNote;
using System.Xml;

namespace OneNoteMDImporter
{
    public static class OneNoteOperator
    {
        public static void CreatePageInSection(string notebookName, string sectionName, Func<string, string> contentXmlBuilder)
        {
            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

            // �m�[�g�u�b�N��ID���擾
            string xml;
            oneNoteApp.GetHierarchy(null, HierarchyScope.hsPages, out xml);

            var doc = new XmlDocument();
            doc.LoadXml(xml);

            // �w�薼�̃m�[�g�u�b�N��ID���擾
            var notebookNode = doc.SelectSingleNode($"//one:Notebook[@name='{notebookName}']", GetNamespaceManager(doc));
            if (notebookNode == null)
            {
                Console.WriteLine($"�m�[�g�u�b�N '{notebookName}' ��������܂���B");
                return;
            }
            string notebookId = notebookNode.Attributes["ID"].Value;

            // �w�薼�̃Z�N�V������ID���擾
            var sectionNode = doc.SelectSingleNode($"//one:Section[@name='{sectionName}']", GetNamespaceManager(doc));
            if (sectionNode == null)
            {
                Console.WriteLine($"�Z�N�V���� '{sectionName}' ��������܂���B");
                return;
            }
            string sectionId = sectionNode.Attributes["ID"].Value;

            // �V�����y�[�W���쐬
            string newPageId;
            oneNoteApp.CreateNewPage(sectionId, out newPageId, NewPageStyle.npsDefault);

            // �R�[���o�b�N�֐��Ńy�[�WXML�𐶐�
            string pageXml = contentXmlBuilder(newPageId);
            oneNoteApp.UpdatePageContent(pageXml, DateTime.MinValue);

            Console.WriteLine("�y�[�W���쐬���܂����B");
        }

        public static XmlNamespaceManager GetNamespaceManager(XmlDocument doc)
        {
            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");
            return nsmgr;
        }
    }
}
