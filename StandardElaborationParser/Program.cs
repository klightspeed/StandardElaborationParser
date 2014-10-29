using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.IO;
using Google.Apis.Drive;
using Google.GData.Spreadsheets;
using TSVCEO.XmlLasdDatabase;

namespace StandardElaborationParser
{
    class WordTableCell
    {
        public int Row;
        public int Col;
        public int RowSpan;
        public int ColSpan;
        public XElement[] Paragraphs;
        public string Text;
        public XElement DocumentXML;

        public override string ToString()
        {
            return String.Format("    <td rowspan=\"{0}\" colspan=\"{1}\">\n      {2}\n    </td>", RowSpan, ColSpan, String.Join("\n      ", Paragraphs.Select(p => p.ToString()).ToArray()));
        }
    }

    class WordTable
    {
        public int Columns { get { return Cells.GetLength(1); } }
        public int Rows { get { return Cells.GetLength(0); } }
        public WordTableCell[,] Cells;

        public override string ToString()
        {
            return String.Format("<table>\n{0}\n</table>",
                String.Join("\n",
                    Enumerable.Range(0, Rows).Select(ri =>
                        String.Format("  <tr>\n{1}\n  </tr>",
                            String.Join("\n",
                                Enumerable.Range(0, Columns).Select(ci =>
                                    Cells[ri, ci] == null ? null : Cells[ri, ci].ToString()
                                ).Where(c => c != null)
                                 .ToArray()
                            )
                        )
                    ).ToArray()
                )
            );
        }
    }

    class WordDocument
    {
        public XElement Body { get; set; }
        public Dictionary<string, XElement> Styles { get; set; }
    }

    class KLADocument
    {
        public List<WordTable> Tables;
        public WordDocument DocumentXML;
        public KeyLearningArea KLA;
    }

    class Program
    {
        static XNamespace ns_xmlPackage = "http://schemas.microsoft.com/office/2006/xmlPackage";
        static XNamespace ns_wpml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static XNamespace ns_lasd = "http://tempuri.org/XmlLasdDatabase.xsd";

        static Dictionary<string, string> YearLevels = new Dictionary<string, string>
        {
            { "prep", "Prep" },
            { "yr1", "Year 1" },
            { "yr2", "Year 2" },
            { "yr3", "Year 3" },
            { "yr4", "Year 4" },
            { "yr5", "Year 5" },
            { "yr6", "Year 6" },
            { "yr7", "Year 7" },
            { "yr8", "Year 8" },
            { "yr9", "Year 9" },
            { "yr10", "Year 10" }
        };

        static Dictionary<string, string> Subjects = new Dictionary<string, string>
        {
            { "eng", "English" },
            { "math", "Mathematics" },
            { "geog", "Geography" },
            { "hist", "History" },
            { "sci", "Science" }
        };

        static void Recurse(XElement node, int depth)
        {
            Console.WriteLine("{0}<{1}>", new String(' ', depth), node.Name.ToString());
            foreach (XElement child in node.Elements())
            {
                Recurse(child, depth + 1);
            }
            Console.WriteLine("{0}</{1}>", new String(' ', depth), node.Name.ToString());
        }

        static List<XElement> CellContent(XElement wordcell, Dictionary<string, XElement> styles)
        {
            List<XElement> output = new List<XElement>();
            XElement list = null;

            foreach (XElement p in wordcell.Elements(ns_wpml + "p"))
            {
                if (p.Value != "")
                {
                    string style = "";
                    XElement ppr = p.Element(ns_wpml + "pPr");
                    XElement pstyle = ppr.Element(ns_wpml + "pStyle");
                    if (pstyle != null)
                    {
                        XAttribute psval = pstyle.Attribute(ns_wpml + "val");
                        style = psval.Value;
                    }

                    if (( p.Elements(ns_wpml + "pPr").SelectMany(e => e.Elements(ns_wpml + "numPr")).SelectMany(e => e.Elements(ns_wpml + "numId")).Count() != 0) ||
                        (styles.ContainsKey(style) && styles[style].Elements(ns_wpml + "pPr").SelectMany(e => e.Elements(ns_wpml + "numPr")).SelectMany(e => e.Elements(ns_wpml + "numId")).Count() != 0))

                    {
                        if (list == null)
                        {
                            list = new XElement(ns_lasd + "ul");
                        }

                        list.Add(new XElement(ns_lasd + "li", p.Value));
                    }
                    else
                    {
                        if (list != null)
                        {
                            output.Add(list);
                            list = null;
                        }

                        output.Add(new XElement(ns_lasd + "p", p.Value));
                    }
                }
            }

            if (list != null)
            {
                output.Add(list);
                list = null;
            }

            return output;
        }

        static string CellText(XElement wordcell)
        {
            List<string> lines = new List<string>();

            foreach (XElement p in wordcell.Elements(ns_wpml + "p"))
            {
                string text = p.Value;
                if (text != "")
                {
                    string style = "";
                    XElement ppr = p.Element(ns_wpml + "pPr");
                    XElement pstyle = ppr.Element(ns_wpml + "pStyle");
                    if (pstyle != null)
                    {
                        XAttribute psval = pstyle.Attribute(ns_wpml + "val");
                        style = psval.Value;
                    }

                    if (style == "Tablebullets")
                    {
                        text = " • " + text;
                    }

                    lines.Add(text);
                }
            }

            return String.Join("\n", lines);
        }

        static WordDocument GetXml(Word.Application app, string docfile)
        {
            string xml;
            
            if (File.Exists(docfile + ".xml"))
            {
                xml = File.ReadAllText(docfile + ".xml");
            }
            else
            {
                Word.Document doc = app.Documents.Open(docfile + ".doc");
                xml = doc.WordOpenXML;
                ((Word._Document)doc).Close(SaveChanges: false);
                File.WriteAllText(docfile + ".xml", xml);
            }

            XElement root = XDocument.Parse(xml).Root;

            return new WordDocument
            {
                Body   = root.Elements(ns_xmlPackage + "part")
                             .SelectMany(e => e.Elements(ns_xmlPackage + "xmlData"))
                             .SelectMany(e => e.Elements(ns_wpml + "document"))
                             .SelectMany(e => e.Elements(ns_wpml + "body"))
                             .Single(),
                Styles = root.Elements(ns_xmlPackage + "part")
                             .SelectMany(e => e.Elements(ns_xmlPackage + "xmlData"))
                             .SelectMany(e => e.Elements(ns_wpml + "styles"))
                             .Single()
                             .Elements(ns_wpml + "style")
                             .ToDictionary(s => s.Attribute(ns_wpml + "styleId").Value, s => s)
            };
        }

        static List<WordTable> GetTables(WordDocument doc)
        {
            List<WordTable> tables = new List<WordTable>();
            int tblno = 0;
            foreach (XElement tbl in doc.Body.Elements(ns_wpml + "tbl"))
            {
                //Console.WriteLine(tbl.ToString());
                WordTable table = new WordTable();
                XElement[] trs = tbl.Elements(ns_wpml + "tr").ToArray();
                int nrcols = tbl.Element(ns_wpml + "tblGrid").Elements(ns_wpml + "gridCol").Count();
                int nrrows = trs.Length;
                table.Cells = new WordTableCell[nrrows, nrcols];
                WordTableCell[] columns = new WordTableCell[nrcols];
                //Console.WriteLine("Table {0}", tblno);
                int trno = 0;
                foreach (XElement tr in tbl.Elements(ns_wpml + "tr"))
                {
                    //Console.WriteLine("  Row {0}", trno);
                    int tcno = 0;
                    foreach (XElement tc in tr.Elements(ns_wpml + "tc"))
                    {
                        //Console.WriteLine("    Col {0}", tcno);
                        XElement tcpr = tc.Element(ns_wpml + "tcPr");
                        XElement vMerge = tcpr.Element(ns_wpml + "vMerge");
                        bool dovMerge = false;

                        if (vMerge != null && vMerge.Attribute(ns_wpml + "val") == null)
                        {
                            dovMerge = true;
                        }

                        XElement gridSpan = tcpr.Element(ns_wpml + "gridSpan");
                        int colspan = 1;

                        if (gridSpan != null)
                        {
                            colspan = Int32.Parse(gridSpan.Attribute(ns_wpml + "val").Value);
                        }

                        if (dovMerge)
                        {
                            columns[tcno].RowSpan++;
                            table.Cells[trno, tcno] = columns[tcno];
                        }
                        else
                        {
                            WordTableCell cell = new WordTableCell
                            {
                                Row = trno,
                                Col = tcno,
                                RowSpan = 1,
                                ColSpan = colspan,
                                Paragraphs = CellContent(tc, doc.Styles).ToArray(),
                                Text = CellText(tc),
                                DocumentXML = tc
                            };
                            columns[tcno] = cell;
                            table.Cells[trno, tcno] = cell;
                        }

                        //Console.WriteLine("      {0}", tcpr.ToString());
                        tcno += colspan;
                    }
                    trno++;
                }
                tables.Add(table);
                tblno++;
            }

            return tables;
        }

        static void Main(string[] args)
        {
            List<KLADocument> kladocs = new List<KLADocument>();
            
            Word.Application app = new Word.Application();
            
            foreach (KeyValuePair<string, string> grade_kvp in YearLevels)
            {
                foreach (KeyValuePair<string, string> subject_kvp in Subjects)
                {
                    string filename = Path.Combine(Environment.CurrentDirectory, @"ac_" + subject_kvp.Key + "_" + grade_kvp.Key + "_se");
                    Console.WriteLine("Reading {0} {1} ({2})", grade_kvp.Value, subject_kvp.Value, filename);
                    KLADocument kla = new KLADocument
                    {
                        KLA = new KeyLearningArea
                        {
                            YearLevelID = grade_kvp.Key,
                            YearLevel = grade_kvp.Value,
                            SubjectID = subject_kvp.Key,
                            Subject = subject_kvp.Value,
                            Groups = new List<AchievementRowGroup>(),
                            Terms = new List<TermDefinition>()
                        },
                        DocumentXML = GetXml(app, filename)
                    };
                    kladocs.Add(kla);
                }
            }

            ((Word._Application)app).Quit();

            foreach (KLADocument kladoc in kladocs)
            {
                Console.WriteLine("Processing {0} {1}", kladoc.KLA.YearLevel, kladoc.KLA.Subject);
                kladoc.Tables = GetTables(kladoc.DocumentXML);

                foreach (WordTable table in kladoc.Tables)
                {
                    int cols = table.Columns;
                    int rows = table.Rows;

                    if (cols >= 7)
                    {
                        //kla.AchievementLevels = Enumerable.Range(0, cols).Select(c => table.Cells[0, c]).Where(c => c != null).Reverse().Take(5).Reverse().Select(c => c.Text).ToList();

                        for (int r = 2; r < rows; r++)
                        {
                            List<WordTableCell> cells = Enumerable.Range(0, cols).Select(c => table.Cells[r, c]).Where(c => c != null).Reverse().ToList();
                            List<string> groups = cells.Skip(5).Reverse().Select(c => c.Text.Replace("\n", ": ")).ToList();
                            WordTableCell[] descs = cells.Take(5).Reverse().ToArray();
                            AchievementRowGroup grp = new AchievementRowGroup
                            {
                                Name = kladoc.KLA.YearLevel + " " + kladoc.KLA.Subject,
                                Groups = kladoc.KLA.Groups,
                                Rows = null,
                                Id = kladoc.KLA.YearLevelID + "::" + kladoc.KLA.SubjectID + "::"
                            };

                            foreach (string grpname in groups)
                            {
                                AchievementRowGroup subgrp = grp.Groups.SingleOrDefault(g => g.Name == grpname);

                                if (subgrp == null)
                                {
                                    subgrp = new AchievementRowGroup
                                    {
                                        Id = grp.Id + (grp.Groups.Count + 1).ToString() + ".",
                                        Name = grpname,
                                        Groups = new List<AchievementRowGroup>(),
                                        Rows = new List<AchievementRow>()
                                    };

                                    grp.Groups.Add(subgrp);
                                }

                                grp = subgrp;
                            }

                            grp.Rows.Add(new AchievementRow
                            {
                                Descriptors = descs.Select(d => new FormattedText { Elements = d.Paragraphs }).ToList(),
                                Id = grp.Id + (grp.Rows.Count + 1).ToString()
                            });
                        }
                    }
                    else if (cols == 2 && table.Cells[0, 0].Text == "Term")
                    {
                        for (int r = 1; r < rows; r++)
                        {
                            List<string> keywords = table.Cells[r, 0].Text.Split(',', ';').Select(k => k.Trim().Replace('\xA0', ' ')).ToList();
                            string name = keywords.FirstOrDefault(k => k.EndsWith("*")) ?? keywords.First();
                            kladoc.KLA.Terms.Add(new TermDefinition
                            {
                                Name = name.TrimEnd('*'),
                                Keywords = keywords.Select(k => k.TrimEnd('*')).ToList(),
                                Description = new FormattedText { Elements = table.Cells[r, 1].Paragraphs }
                            });
                        }
                    }
                }
            }

            foreach (KLADocument kladoc in kladocs)
            {
                kladoc.KLA.ToXDocument().Save(String.Format("{0}-{1}.xml", kladoc.KLA.YearLevelID, kladoc.KLA.SubjectID));
            }

            //Recurse(xdoc.Root, 0);

        }
    }
}
