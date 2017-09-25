using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

using Microsoft.Office.Interop.OneNote;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OneNoteReader
{
    public class OneNoteNode
    {
        public string Id { get; set; }

        public string Title { get; set; }

        public string Type { get; set; }

        public string Url { get; set; }
    }

    public class NotebookInfo : OneNoteNode
    {
        public string Path { get; set; }

        [JsonProperty(Order = 1)]
        public SectionBase[] Sections { get; set; }
    }

    public class SectionBase : OneNoteNode
    {
        public string Path { get; set; }
    }

    public class SectionGroupInfo : SectionBase
    {
        [JsonProperty(Order = 1)]
        public SectionBase[] Sections { get; set; }
    }

    public class SectionInfo : SectionBase
    {
        [JsonProperty(Order = 1)]
        public List<PageInfo> Pages { get; } = new List<PageInfo>();
    }

    public class PageInfo : OneNoteNode
    {
        [JsonProperty(Order = 1)]
        public List<PageInfo> Pages { get; } = new List<PageInfo>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            var oneNoteApp = new Application();

            var notebooks = oneNoteApp.GetNotebooks();

            var settings = new JsonSerializerSettings()
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver()
            };
            var json = JsonConvert.SerializeObject(notebooks, settings);
            System.IO.File.WriteAllText("notebooks.json", json);
        }
    }
}
