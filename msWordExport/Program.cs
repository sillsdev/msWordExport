// ---------------------------------------------------------------------------------------------
#region // Copyright (c) 2013, SIL International. All Rights Reserved.
// <copyright from='2013' to='2013' company='SIL International'>
//		Copyright (c) 2013, SIL International. All Rights Reserved.
//
//		Distributable under the terms of either the Common Public License or the
//		GNU Lesser General Public License, as specified in the LICENSING.txt file.
// </copyright>
#endregion
//
// File: program.cs
// Responsibility: Greg Trihus
// ---------------------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Xsl;
using Mono.Options;


namespace msWordExport
{
    class Program
    {
        static int _verbosity = 0;

        static void Main(string[] args)
        {
            var xsltSettings = new XsltSettings() { EnableDocumentFunction = true };
            bool showHelp = false;
            var myArgs = new List<string>();
            var output = "main.doc";
            var title = "My Title";
            // see: http://stackoverflow.com/questions/491595/best-way-to-parse-command-line-arguments-in-c
            var p = new OptionSet {
                { "o|output=", "the {OUTPUT} name.\n" +
                        "Defaults to main.",
                   v => output = v },
                { "t|title=", "the {TITLE}.\n" +
                        "Defaults to My Title.",
                   v => output = v },
                { "d|define=", "define argument:value for transformation.",
                   v => myArgs.Add (v) },
                { "v", "increase debug message verbosity",
                   v => { if (v != null) ++_verbosity; } },
                { "h|help",  "show this message and exit", 
                   v => showHelp = v != null },
            };

            List<string> extra;
            try
            {
                extra = p.Parse(args);
                if (extra.Count == 0)
                {
                    Console.WriteLine("Enter full file name to process");
                    extra.Add(Console.ReadLine());
                }
            }
            catch (OptionException e)
            {
                Console.Write("msWordExport: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `msWordExport --help' for more information.");
                return;
            }

            if (showHelp || extra.Count != 1)
            {
                ShowHelp(p);
                return;
            }

            if (Path.GetDirectoryName(extra[0]) != "" && Path.GetDirectoryName(output) == "")
            {
                output = Path.Combine(Path.GetDirectoryName(extra[0]), output);
            }

            var xsltArgs = new XsltArgumentList();
            CreateArgumentList(myArgs, xsltArgs);

            var cssData = GetData(extra[0]);
            var xDoc = new XmlDocument {XmlResolver = null};
            GetXhtml(extra, xDoc);
            ApplyRules(xDoc, cssData);
            var nsmgr = GetNamespaceManager(xDoc);
            InsertStyles(cssData, xDoc, nsmgr);
            InsertTitle(xDoc, nsmgr, title);
            UpdateImgSrc(xDoc, nsmgr);
            WriteXhtml(output, xDoc);
            Process.Start(output);
        }

        private static void UpdateImgSrc(XmlDocument xDoc, XmlNamespaceManager nsmgr)
        {
            var imgBase = xDoc.SelectSingleNode("//xhtml:meta[@name = 'linkedFilesRootDir']/@content", nsmgr).Value;
            foreach (XmlNode node in xDoc.SelectNodes("//xhtml:img/@src", nsmgr))
            {
                if (!Path.IsPathRooted(node.InnerText))
                {
                    node.InnerText = Path.Combine(imgBase, node.InnerText);
                }
            }
        }

        private static void WriteXhtml(string output, XmlDocument xDoc)
        {
            var settings = new XmlWriterSettings
                {
                    OmitXmlDeclaration = true,
                    Indent = true,
                    Encoding = new UTF8Encoding(true)
                };
            var xWriter = XmlWriter.Create(output, settings);
            xDoc.Save(xWriter);
            xWriter.Close();
        }

        private static void InsertTitle(XmlDocument xDoc, XmlNamespaceManager nsmgr, string title)
        {
            var titleNode = xDoc.SelectSingleNode("//xhtml:head/xhtml:title", nsmgr);
            titleNode.InnerText = title;
        }

        private static void InsertStyles(string cssData, XmlDocument xDoc, XmlNamespaceManager nsmgr)
        {
            var style = new StringBuilder("\n");
            var declPattern = new Regex("\\.([-a-zA-Z]+)\\s{");
            var propPattern = new Regex("\\s+([-a-zA-Z]+)\\:");
            var isSpan = false;
            var firstDecl = false;
            var propertyList = new List<string>(5)
                {
                    "padding-left",
                    "text-indent",
                    "margin-left",
                    "padding-bottom",
                    "padding-top"
                };
            foreach (string line in cssData.Split(new[] {'\n'}))
            {
                if (line.IndexOf(":before") == -1 && line.IndexOf(":after") == -1)
                {
                    var matchDecl = declPattern.Match(line);
                    if (matchDecl.Success)
                    {
                        firstDecl = true;
                        var xpath = string.Format("//*[@class='{0}']", matchDecl.Groups[1].Value);
                        var node = (XmlElement) xDoc.SelectSingleNode(xpath);
                        isSpan = (node != null && node.LocalName == "span");
                    }
                    else if (isSpan)
                    {
                        var matchProp = propPattern.Match(line);
                        if (matchProp.Success)
                        {
                            if (propertyList.Contains(matchProp.Groups[1].Value))
                            {
                                continue;
                            }
                        }
                    }
                    if (firstDecl)
                    {
                        style.Append(line + "\n");
                    }
                }
            }
            var headNode = xDoc.SelectSingleNode("//xhtml:head", nsmgr);
            var styleNode = xDoc.CreateElement("style", "http://www.w3.org/1999/xhtml");
            styleNode.InnerText = style.ToString();
            headNode.AppendChild(styleNode);
            var styleLink = xDoc.SelectSingleNode("//xhtml:link[@href]", nsmgr);
            headNode.RemoveChild(styleLink);
        }

        private static void GetXhtml(List<string> extra, XmlDocument xDoc)
        {
            var xfs = new StreamReader(extra[0]);
            xDoc.Load(xfs);
            xfs.Close();
        }

        private static void UpdateStyleLink(XmlDocument xDoc, XmlNamespaceManager nsmgr, string cssName)
        {
            var styleLinkNode = xDoc.SelectSingleNode("//xhtml:link[@rel='stylesheet']", nsmgr);
            styleLinkNode.Attributes["href"].InnerText = cssName;
        }

        private static void ApplyRules(XmlDocument xDoc, string cssData)
        {
            foreach (XmlNode node in xDoc.SelectNodes("//*[@class]"))
            {
                var className = node.Attributes["class"].InnerText;
                ApplyBeforeRule(cssData, className, node);
                ApplyBetweenRules(cssData, className, node);
                ApplyAfterRules(cssData, className, node);
            }
        }

        private static void ApplyAfterRules(string cssData, string className, XmlNode node)
        {
            var afterRules = CollectRules(cssData, "\n\\.([-a-zA-Z]+):after\\s{\\scontent:\\s\"([^\"]*)");
            if (afterRules.ContainsKey(className))
            {
                node.InnerText = node.InnerText + afterRules[className];
            }
        }

        private static void ApplyBetweenRules(string cssData, string className, XmlNode node)
        {
            var betweenRules = CollectRules(cssData, "\n\\.([-a-zA-Z]+)>\\.xitem\\s\\+\\s\\.xitem:before\\s{\\scontent:\\s\"([^\"]*)");
            if (betweenRules.ContainsKey(className))
            {
                var first = true;
                foreach (XmlElement child in node.SelectNodes("*"))
                {
                    if (first)
                    {
                        if (child.HasAttribute("class") && child.Attributes["class"].Value != "xitem")
                        {
                            continue;
                        }
                        first = false;
                        continue;
                    }
                    child.InnerText = betweenRules[className] + child.InnerText;
                }
            }
        }

        private static void ApplyBeforeRule(string cssData, string className, XmlNode node)
        {
            var beforeRules = CollectRules(cssData, "\n\\.([-a-zA-Z]+):before\\s{\\scontent:\\s\"([^\"]*)");
            if (beforeRules.ContainsKey(className))
            {
                node.InnerText = beforeRules[className] + node.InnerText;
            }
        }

        private static string GetData(string fullName)
        {
            var name = Path.GetFileNameWithoutExtension(fullName);
            var folder = Path.GetDirectoryName(fullName);
            var cssName = name + ".css";
            var cssFullName = Path.Combine(folder, cssName);
            var fs = new StreamReader(cssFullName);
            var cssData = fs.ReadToEnd();
            fs.Close();
            return cssData;
        }

        private static Dictionary<string, string> CollectRules(string cssData, string patternDesc)
        {
            var pattern = new Regex(patternDesc);
            var rules = new Dictionary<string, string>();
            foreach (Match match in pattern.Matches(cssData))
            {
                rules[match.Groups[1].Value] = match.Groups[2].Value;
            }
            return rules;
        }

        private static void CreateArgumentList(IEnumerable<string> myArgs, XsltArgumentList xsltArgs)
        {
            foreach (string definition in myArgs)
            {
                if (definition.Contains(":"))
                {
                    var defParse = definition.Split(':');
                    xsltArgs.AddParam(defParse[0], "", defParse[1]);
                }
                else
                {
                    xsltArgs.AddParam(definition, "", true);
                }
            }
        }

        protected static XmlNamespaceManager GetNamespaceManager(XmlDocument xmlDocument, string defaultNs)
        {
            var root = xmlDocument.DocumentElement;
            var nsManager = new XmlNamespaceManager(xmlDocument.NameTable);
            foreach (XmlAttribute attribute in root.Attributes)
            {
                if (attribute.Name == "xmlns")
                {
                    nsManager.AddNamespace(defaultNs, attribute.Value);
                    continue;
                }
                var namePart = attribute.Name.Split(':');
                if (namePart[0] == "xmlns")
                    nsManager.AddNamespace(namePart[1], attribute.Value);
            }
            return nsManager;
        }

        protected static XmlNamespaceManager GetNamespaceManager(XmlDocument xmlDocument)
        {
            return GetNamespaceManager(xmlDocument, "xhtml");
        }

        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage: msWordExport [OPTIONS]+ xhtmlFile");
            Console.WriteLine("Import xhtml file into ms Word.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        static void Debug(string format, params object[] args)
        {
            if (_verbosity > 0)
            {
                Console.Write("# ");
                Console.WriteLine(format, args);
            }
        }
    }
}
