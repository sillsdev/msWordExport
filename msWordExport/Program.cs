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
// License: LGPL2
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
            bool showHelp = false;
            var output = "main.doc";
            var title = "My Title";
            // see: http://stackoverflow.com/questions/491595/best-way-to-parse-command-line-arguments-in-c
            var p = new OptionSet {
                { "o|output=", "the {OUTPUT} name.\n" +
                        "Defaults to main.",
                   v => output = v },
                { "t|title=", "the {TITLE}.\n" +
                        "Defaults to My Title.",
                   v => title = v },
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

            var cssData = GetData(extra[0]);
            var xDoc = new XmlDocument {XmlResolver = null};
            GetXhtml(extra, xDoc);
            ApplyRules(xDoc, cssData);
            var nsmgr = GetNamespaceManager(xDoc);
            ChangeDiv2P("//xhtml:div[xhtml:span]", xDoc, nsmgr);
            ChangeDiv2P("//xhtml:div[@class='letter']", xDoc, nsmgr);
            InsertStyles(cssData, xDoc, nsmgr);
            InsertTitle(xDoc, nsmgr, title);
            UpdateImgSrc(xDoc, nsmgr);
            WriteXhtml(output, xDoc);
            Process.Start(output);
        }

        private static void ChangeDiv2P(string xpath, XmlDocument xDoc, XmlNamespaceManager nsmgr)
        {
            var divNode = xDoc.SelectNodes(xpath, nsmgr);
            for (int divIdx = 0; divIdx < divNode.Count; divIdx += 1)
            {
                var divElement = divNode[divIdx];
                var pNode = xDoc.CreateElement("p", "http://www.w3.org/1999/xhtml");
                var attrNode = new XmlAttribute[divElement.Attributes.Count];
                divElement.Attributes.CopyTo(attrNode,0);
                for (int attrIdx = 0; attrIdx < attrNode.Length; attrIdx += 1)
                {
                    pNode.Attributes.Append(attrNode[attrIdx]);
                }
                foreach (XmlNode element in divElement.ChildNodes)
                {
                    pNode.AppendChild(element.CloneNode(true));
                }
                var parent = divElement.ParentNode;
                parent.InsertAfter(pNode, divElement.PreviousSibling);
                parent.RemoveChild(divElement);
            }
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
            var propertyList = new List<string>()
                {
                    "padding-left",
                    "text-indent",
                    "margin-left",
                    "padding-bottom",
                    "padding-top",
                    "border-style",
                    "border-top",
                    "border-bottom",
                    "border-left",
                    "border-right",
                    "border-width",
                    "border-color"
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

        private static void ApplyRules(XmlDocument xDoc, string cssData)
        {
            var beforeRules = CollectRules(cssData, "\n\\.([-a-zA-Z]+):before\\s{\\scontent:\\s\"([^\"]*)");
            var betweenRules = CollectRules(cssData, "\n\\.([-a-zA-Z]+)>\\.xitem\\s\\+\\s\\.xitem:before\\s{\\scontent:\\s\"([^\"]*)");
            var afterRules = CollectRules(cssData, "\n\\.([-a-zA-Z]+):after\\s{\\scontent:\\s\"([^\"]*)");
            foreach (XmlNode node in xDoc.SelectNodes("//*[@class]"))
            {
                var classNames = node.Attributes["class"].InnerText;
                var classList = classNames.Split(new[] {' '});
                var nonContentClass = string.Empty;
                foreach (var className in classList)
                {
                    var found = ApplyBeforeRule(beforeRules, className, node);
                    found |= ApplyBetweenRules(betweenRules, className, node);
                    found |= ApplyAfterRules(afterRules, className, node);
                    if (!found)
                    {
                        if (nonContentClass != string.Empty)
                        {
                            throw new ArgumentException(string.Format("Two non-content classes {0} and {1}", nonContentClass, className));
                        }
                        nonContentClass = className;
                    }
                }
                if (classList.Length > 1)
                {
                    node.Attributes["class"].InnerText = nonContentClass;
                }

            }
        }

        private static bool ApplyBeforeRule(IDictionary<string, string> beforeRules, string className, XmlNode node)
        {
            var found = beforeRules.ContainsKey(className);
            if (found)
            {
                node.InnerText = beforeRules[className] + node.InnerText;
            }
            return found;
        }

        private static bool ApplyBetweenRules(IDictionary<string, string> betweenRules, string className, XmlNode node)
        {
            var found = betweenRules.ContainsKey(className);
            if (found)
            {

                var first = true;
                foreach (XmlElement child in node.SelectNodes("./*[@class = 'xitem']"))
                {
                    if (first)
                    {
                        first = false;
                        continue;
                    }
                    // Although Flex outputs in between rule, it also inserts in between punctuation
                    // so here we make sure the rule should be applied
                    var prevClass = child.PreviousSibling.SelectSingleNode("@class");
                    if (prevClass == null || prevClass.InnerText != "xitem")
                    {
                        continue;
                    }
                    child.InnerText = betweenRules[className] + child.InnerText;
                }
            }
            return found;
        }

        private static bool ApplyAfterRules(IDictionary<string, string> afterRules, string className, XmlNode node)
        {
            var found = afterRules.ContainsKey(className);
            if (found)
            {
                node.InnerText = node.InnerText + afterRules[className];
            }
            return found;
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
