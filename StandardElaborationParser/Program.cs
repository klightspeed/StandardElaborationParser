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

namespace StandardElaborationParser
{
    class TableCell
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

    class Table
    {
        public int Columns { get { return Cells.GetLength(1); } }
        public int Rows { get { return Cells.GetLength(0); } }
        public TableCell[,] Cells;

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

    class Document
    {
        public XElement Body { get; set; }
        public Dictionary<string, XElement> Styles { get; set; }
    }

    class Group
    {
        protected static readonly XNamespace ns = "http://tempuri.org/XmlLasdDatabase.xsd";
        public Dictionary<string, Group> Groups = new Dictionary<string, Group>();
        public List<TableCell[]> Descriptors = new List<TableCell[]>();

        public IEnumerable<KLARow> GetDescriptors(string id)
        {
            int index = 1;

            foreach (Group g in Groups.Values)
            {
                string gid = (id == null ? "" : (id + ".")) + index.ToString();
                foreach (KLARow t in g.GetDescriptors(gid))
                {
                    yield return t;
                }

                index++;
            }

            foreach (TableCell[] row in Descriptors)
            {
                yield return new KLARow { GroupID = id, AchievementDescriptors = row };
            }
        }

        public IEnumerable<XElement> ToXML()
        {
            foreach (KeyValuePair<string, Group> g_kvp in Groups)
            {
                yield return new XElement(ns + "group",
                    new XAttribute("name", g_kvp.Key),
                    g_kvp.Value.ToXML()
                );
            }

            foreach (TableCell[] row in Descriptors)
            {
                yield return new XElement(ns + "row",
                    row.Select(d => new XElement(ns + "descriptor", d.Paragraphs))
                );
            }
        }
    }

    class KLARow
    {
        public string GroupID { get; set; }
        public TableCell[] AchievementDescriptors { get; set; }
    }

    class KLA
    {
        protected static readonly XNamespace ns = "http://tempuri.org/XmlLasdDatabase.xsd";
        public string YearLevel;
        public string Subject;
        public string YearLevelID;
        public string SubjectID;
        public List<Tuple<string, TableCell>> Definitions = new List<Tuple<string,TableCell>>();
        public List<string> AchievementLevels = new List<string>();
        public Group RootGroup = new Group();
        public List<Table> Tables;
        public Document DocumentXML;
        public IEnumerable<KLARow> Rows
        {
            get
            {
                return RootGroup.GetDescriptors(null);
            }
        }

        public XDocument ToXDocument()
        {
            return new XDocument(
                new XElement(ns + "kla",
                    new XAttribute("yearLevel", YearLevel),
                    new XAttribute("yearLevelId", YearLevelID),
                    new XAttribute("subject", Subject),
                    new XAttribute("subjectId", SubjectID),
                    new XElement(ns + "terms",
                        Definitions.Select(t => new XElement(ns + "term",
                            t.Item1.Split(';', ',').Select(k => new XElement(ns + "keyword", k.Trim())),
                            new XElement(ns + "description", t.Item2.Paragraphs)
                        ))
                    ),
                    RootGroup.ToXML()
                )
            );
        }
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

        static Document GetXml(Word.Application app, string docfile)
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

            return new Document
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

        static List<Table> GetTables(Document doc)
        {
            List<Table> tables = new List<Table>();
            int tblno = 0;
            foreach (XElement tbl in doc.Body.Elements(ns_wpml + "tbl"))
            {
                //Console.WriteLine(tbl.ToString());
                Table table = new Table();
                XElement[] trs = tbl.Elements(ns_wpml + "tr").ToArray();
                int nrcols = tbl.Element(ns_wpml + "tblGrid").Elements(ns_wpml + "gridCol").Count();
                int nrrows = trs.Length;
                table.Cells = new TableCell[nrrows, nrcols];
                TableCell[] columns = new TableCell[nrcols];
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
                            TableCell cell = new TableCell
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
            List<KLA> klas = new List<KLA>();
            
            Word.Application app = new Word.Application();
            
            foreach (KeyValuePair<string, string> grade_kvp in YearLevels)
            {
                foreach (KeyValuePair<string, string> subject_kvp in Subjects)
                {
                    string filename = Path.Combine(Environment.CurrentDirectory, @"ac_" + subject_kvp.Key + "_" + grade_kvp.Key + "_se");
                    Console.WriteLine("Reading {0} {1} ({2})", grade_kvp.Value, subject_kvp.Value, filename);
                    KLA kla = new KLA
                    {
                        YearLevelID = grade_kvp.Key,
                        YearLevel = grade_kvp.Value,
                        SubjectID = subject_kvp.Key,
                        Subject = subject_kvp.Value,
                        DocumentXML = GetXml(app, filename)
                    };
                    klas.Add(kla);
                }
            }

            ((Word._Application)app).Quit();

            foreach (KLA kla in klas)
            {
                Console.WriteLine("Processing {0} {1}", kla.YearLevel, kla.Subject);
                kla.Tables = GetTables(kla.DocumentXML);

                foreach (Table table in kla.Tables)
                {
                    int cols = table.Columns;
                    int rows = table.Rows;

                    if (cols >= 7)
                    {
                        kla.AchievementLevels = Enumerable.Range(0, cols).Select(c => table.Cells[0, c]).Where(c => c != null).Reverse().Take(5).Reverse().Select(c => c.Text).ToList();

                        for (int r = 2; r < rows; r++)
                        {
                            List<TableCell> cells = Enumerable.Range(0, cols).Select(c => table.Cells[r, c]).Where(c => c != null).Reverse().ToList();
                            List<string> groups = cells.Skip(5).Reverse().Select(c => c.Text.Replace("\n", ": ")).ToList();
                            TableCell[] descs = cells.Take(5).Reverse().ToArray();
                            Group grp = kla.RootGroup;

                            foreach (string grpname in groups)
                            {
                                if (!grp.Groups.ContainsKey(grpname))
                                {
                                    grp.Groups[grpname] = new Group();
                                }

                                grp = grp.Groups[grpname];
                            }

                            grp.Descriptors.Add(descs);
                        }
                    }
                    else if (cols == 2 && table.Cells[0, 0].Text == "Term")
                    {
                        for (int r = 1; r < rows; r++)
                        {
                            kla.Definitions.Add(new Tuple<string, TableCell>(table.Cells[r, 0].Text, table.Cells[r, 1]));
                        }
                    }
                }
            }

            foreach (KLA kla in klas)
            {
                kla.ToXDocument().Save(String.Format("{0}-{1}.xml", kla.YearLevelID, kla.SubjectID));
            }

            //Recurse(xdoc.Root, 0);

        }
    }
}
