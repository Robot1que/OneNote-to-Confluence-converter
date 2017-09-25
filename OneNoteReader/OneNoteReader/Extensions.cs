using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

using Microsoft.Office.Interop.OneNote;

namespace OneNoteReader
{
    public static class Extensions
    {

        public static string ToAscii(this string value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            var asciiChars = value.Where(ch => Encoding.UTF8.GetByteCount(new char[] { ch }) == 1).ToArray();
            return new String(asciiChars);
        }

        public static NotebookInfo[] GetNotebooks(this Application oneNoteApp)
        {
            oneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out string xml);
            var doc = XDocument.Parse(xml);
            var ns = doc.Root.Name.Namespace;
            var elements = doc.Descendants(ns + "Notebook").ToArray();

            PrintWrongNames(elements);

            var notebookInfos = new List<NotebookInfo>(elements.Length);

            foreach (var element in elements)
            {
                var notebook =
                    new NotebookInfo()
                    {
                        Id = element.Attribute("ID").Value,
                        Title = element.Attribute("name").Value.ToAscii(),
                        Path = element.Attribute("path").Value,
                        Type = element.Name.LocalName
                    };

                oneNoteApp.GetHyperlinkToObject(notebook.Id, "", out string url);

                notebook.Sections = oneNoteApp.GetSections(notebook.Id);
                notebook.Url = url;

                Console.WriteLine(notebook.Title);

                notebookInfos.Add(notebook);
            }

            return notebookInfos.ToArray();
        }

        public static SectionBase[] GetSections(this Application oneNoteApp, string notebookId)
        {
            oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out string xml);
            var doc = XDocument.Parse(xml);
            var ns = doc.Root.Name.Namespace;

            var notebook = doc.Elements().First();

            return oneNoteApp.GetSections(notebook);
        }

        private static SectionBase[] GetSections(this Application oneNoteApp, XElement root)
        {
            var elements =
                root.Elements()
                    .Where(
                        element =>
                            element.Attribute("isRecycleBin") == null &&
                            element.Attribute("isDeletedPages") == null
                    )
                    .ToArray();

            PrintWrongNames(elements);

            var sections = new List<SectionBase>(elements.Length);

            foreach (var element in elements)
            {
                var sectionId = element.Attribute("ID").Value;
                SectionBase section = null;

                if (element.Name.LocalName == "SectionGroup")
                {
                    section = new SectionGroupInfo()
                    {
                        Sections = oneNoteApp.GetSections(element)
                    };
                }
                else if (element.Name.LocalName == "Section")
                {
                    var sectionInfo = new SectionInfo();
                    section = sectionInfo;
                    oneNoteApp.GetPages(sectionInfo.Pages, 1, sectionId);
                }

                if (section != null)
                {
                    oneNoteApp.GetHyperlinkToObject(sectionId, "", out string url);

                    section.Id = sectionId;
                    section.Title = element.Attribute("name").Value.ToAscii();
                    section.Path = element.Attribute("path").Value;
                    section.Type = element.Name.LocalName;
                    section.Url = url;

                    Console.WriteLine("    " + section.Title);

                    sections.Add(section);
                }
            }

            return sections.ToArray();
        }

        public static void GetPages(
            this Application oneNoteApp,
            List<PageInfo> pages,
            int currentPageLevel,
            string sectionId)
        {
            oneNoteApp.GetHierarchy(sectionId, HierarchyScope.hsPages, out string xml);
            var doc = XDocument.Parse(xml);
            var ns = doc.Root.Name.Namespace;
            var elements =
                doc.Descendants(ns + "Page")
                    .Where(
                        element =>
                            element.Attribute("isRecycleBin") == null &&
                            element.Attribute("isDeletedPages") == null
                    )
                    .ToArray();

            PrintWrongNames(elements);

            oneNoteApp.AddPages(elements, pages);
        }

        public static void AddPages(
            this Application oneNoteApp,
            XElement[] elements,
            List<PageInfo> pages)
        {
            var containerStack = new Stack<List<PageInfo>>();
            containerStack.Push(pages);
            var currentPageLevel = 1;
            foreach (var element in elements)
            {
                var pageId = element.Attribute("ID").Value;
                var pageLevel = int.Parse(element.Attribute("pageLevel").Value);

                oneNoteApp.GetHyperlinkToObject(pageId, "", out string url);

                var page =
                    new PageInfo()
                    {
                        Id = pageId,
                        Title = element.Attribute("name").Value.ToAscii(),
                        Type = element.Name.LocalName,
                        Url = url
                    };

                if (pageLevel > currentPageLevel)
                {
                    currentPageLevel = pageLevel;

                    if (containerStack.Any())
                    {
                        var pageList = containerStack.Peek();
                        if (pageList.Any())
                        {
                            containerStack.Push(pageList.Last().Pages);
                        }
                    }
                }
                else if (pageLevel < currentPageLevel)
                {
                    currentPageLevel = pageLevel;
                    if (containerStack.Count > 1)
                    {
                        containerStack.Pop();
                    }
                }

                containerStack.Peek().Add(page);
            }
        }

        private static void PrintWrongNames(XElement[] elements)
        {
            var names =
                elements
                    .Select(element => element.Attribute("name").Value)
                    .Where(name => name.Any(ch => Encoding.UTF8.GetByteCount(new char[] { ch }) != 1))
                    .ToArray();

            System.IO.File.AppendAllLines("wrongnames.txt", names);
        }
    }
}
