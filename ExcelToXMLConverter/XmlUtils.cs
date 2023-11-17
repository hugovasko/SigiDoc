using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ExcelToXMLConverter
{
    internal static class XmlUtils
    {
        public static string Prettify(string xml)
        {
            var stringBuilder = new StringBuilder();

            var element = XElement.Parse(xml);

            var settings = new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                Indent = true,
                NewLineOnAttributes = false
            };

            using (var xmlWriter = XmlWriter.Create(stringBuilder, settings))
            {
                element.Save(xmlWriter);
            }

            return stringBuilder.ToString();
        }

        public static void UpdateSealsList(XNamespace ns, string filename, string sequence, XDocument sealsList)
        {
            var listElement = sealsList.Descendants(ns + "list").FirstOrDefault();

            if (listElement != null)
            {
                bool isItemInList = listElement.Descendants(ns + "item").Any(i => i.Attribute("n")?.Value == filename);

                if (!isItemInList)
                {
                    var newListItem = new XElement(ns + "item");
                    newListItem.SetAttributeValue("n", filename);
                    newListItem.SetAttributeValue("sortKey", sequence);
                    listElement.Add(newListItem);

                    var sortedListItems = listElement.Descendants(ns + "item").OrderBy(i => i.Attribute("sortKey")?.Value).ToList();

                    listElement.ReplaceNodes(sortedListItems);

                    sealsList.Save(@"../all_seals.xml");

                    sealsList = XDocument.Load(@"../all_seals.xml");
                }
            }
        }

        public static void ProcessInterpretiveOrDiplomaticText(XDocument xmlTemplate, XNamespace ns, Dictionary<string, string> allValues, string header, string type, string subtype, string partN)
        {
            var text = allValues.GetValueOrDefault(header);

            if (string.IsNullOrEmpty(text))
            {
                Console.WriteLine($"{type} text is missing or empty.");
                return;
            }

            var separator = "--";
            var splitParts = text.Split(new[] { separator }, StringSplitOptions.None);

            if (splitParts.Length != 2)
            {
                Console.WriteLine($"Invalid {type} text format. Could not find separator '--'.");
                return;
            }

            var obvText = splitParts[0].Trim();
            var revText = splitParts[1].Trim();

            var obvLines = obvText.Split('/');
            var revLines = revText.Split('/');

            var editionElement = xmlTemplate.Descendants(ns + "div").FirstOrDefault(e => (string)e.Attribute("type") == type && (string)e.Attribute("subtype") == subtype);

            if (editionElement == null)
            {
                Console.WriteLine($"Could not find the expected '{type}' div in the XML template.");
                return;
            }

            var obvElement = editionElement.Descendants().FirstOrDefault(e => e.Name == ns + "div" && (string)e.Attribute("type") == "textpart" && (string)e.Attribute("n") == partN);

            if (obvElement == null)
            {
                Console.WriteLine($"Could not find the expected 'textpart' div with n='{partN}' in the XML template.");
                return;
            }

            var revElement = editionElement.Descendants().FirstOrDefault(e => e.Name == ns + "div" && (string)e.Attribute("type") == "textpart" && (string)e.Attribute("n") == "rev");

            if (revElement == null)
            {
                Console.WriteLine($"Could not find the expected 'textpart' div with n='rev' in the XML template.");
                return;
            }

            obvElement.Elements(ns + "ab").Remove();
            revElement.Elements(ns + "ab").Remove();

            foreach (var line in obvLines)
            {
                var abElement = new XElement(ns + "ab", line.Trim());
                obvElement.Add("\n");
                obvElement.Add(abElement);
            }

            foreach (var line in revLines)
            {
                var abElement = new XElement(ns + "ab", line.Trim());
                revElement.Add("\n");
                revElement.Add(abElement);
            }

            obvElement.Add("\n");
            revElement.Add("\n");
        }
    }


}