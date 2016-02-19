using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.IO;
using TSVCEO.XmlLasdDatabase;
using TSVCEO.OOXML.Packaging;
using Ionic.Zip;
using System.Net;

namespace StandardElaborationParser
{
    public class xmlns
    {
        public static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace dcterms = "http://purl.org/dc/terms/";
        public static readonly XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    }

    public class WordTableCell
    {
        public int Row;
        public int Col;
        public int RowSpan;
        public int ColSpan;
        public XElement[] Paragraphs;
        public string Text
        {
            get
            {
                List<string> lines = new List<string>();

                foreach (XElement para in Paragraphs)
                {
                    if (para.Name.LocalName == "p")
                    {
                        lines.Add(para.Value);
                    }
                    else if (para.Name.LocalName == "ul")
                    {
                        foreach (string line in GetListParagraphLines(para.Elements(), 1))
                        {
                            lines.Add(line);
                        }
                    }
                }

                return String.Join("\n", lines);
            }
        }

        public override string ToString()
        {
            return String.Format("    <td rowspan=\"{0}\" colspan=\"{1}\">\n      {2}\n    </td>", RowSpan, ColSpan, String.Join("\n      ", Paragraphs.Select(p => p.ToString()).ToArray()));
        }

        protected IEnumerable<string> GetListParagraphLines(IEnumerable<XElement> paragraphs, int listdepth)
        {
            foreach (XElement para in paragraphs)
            {
                if (para.Name.LocalName == "li")
                {
                    yield return new String(' ', listdepth * 2) + " • " + para.Value;
                }
                else if (para.Name.LocalName == "ul")
                {
                    foreach (string line in GetListParagraphLines(para.Elements(), listdepth + 1))
                    {
                        yield return line;
                    }
                }
            }
        }
    }

    public class WordTable
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

    public class Program
    {
        private static XNamespace ns_lasd = "http://gtmj.tsv.catholic.edu.au/XmlLasdDatabase.xsd";

        private static Dictionary<string, string> YearLevels;

        private static Dictionary<string, string> Subjects;

        private static Dictionary<string, Dictionary<string, string>> AchievementLevels;

        private static Dictionary<string, string> GradeAchievementLevelRefs;

        private static Dictionary<string, Dictionary<string, string[]>> YearLevelGroupings;

        private static Dictionary<string, string[]> SubjectGroupings;

        private static void ProcessConfig()
        {
            XDocument configfile = XDocument.Load(File.Open("stdelabs.xml", FileMode.Open));
            XElement root = configfile.Root;
            YearLevels = 
                root.Element("YearLevels")
                    .Elements("YearLevel")
                    .ToDictionary(
                        e => e.Attribute("id").Value, 
                        e => e.Value
                    );
            Subjects = 
                root.Element("Subjects")
                    .Elements("Subject")
                    .ToDictionary(
                        e => e.Attribute("id").Value, 
                        e => e.Value
                    );
            AchievementLevels = 
                root.Element("YearLevelGroups")
                    .Elements("YearLevelGroup")
                    .ToDictionary(
                        e => e.Attribute("id").Value,
                        e => e.Elements("AchievementLevel")
                              .ToDictionary(
                                al => al.Attribute("id").Value,
                                al => al.Value
                              )
                    );
            GradeAchievementLevelRefs =
                root.Element("YearLevelGroups")
                    .Elements("YearLevelGroup")
                    .SelectMany(
                        e => e.Elements("YearLevel")
                              .Select(yl => new { grp = e.Attribute("id").Value, level = yl.Attribute("id").Value })
                    )
                    .ToDictionary(g => g.level, g => g.grp);
            YearLevelGroupings =
                root.Element("SubjectGroups")
                    .Elements("SubjectGroup")
                    .ToDictionary(
                        e => e.Attribute("id").Value,
                        e => e.Elements("YearLevelGroup")
                              .Select(yg => new { grp = yg.Attribute("id").Value, levels = yg.Elements("YearLevel").Select(yl => yl.Attribute("id").Value).ToArray() })
                              .Union(e.Elements("YearLevel").Select(yl => yl.Attribute("id").Value).Select(yl => new { grp = yl, levels = new[] { yl } }))
                              .ToDictionary(g => g.grp, g => g.levels)
                    );
            SubjectGroupings =
                root.Element("SubjectGroups")
                    .Elements("SubjectGroup")
                    .ToDictionary(
                        e => e.Attribute("id").Value,
                        e => e.Elements("Subject")
                              .Select(s => s.Attribute("id").Value)
                              .ToArray()
                    );
        }

        private static IEnumerable<XNode> ParagraphContent(XElement para, Dictionary<string, XElement> styles)
        {
            foreach (XElement el in para.Elements())
            {
                if (el.Name == xmlns.w + "r")
                {
                    foreach (XElement text in el.Elements(xmlns.w + "t"))
                    {
                        yield return new XText(text.Value);
                    }
                }
                else if (el.Name == xmlns.w + "hyperlink")
                {
                    foreach (XNode node in ParagraphContent(el, styles))
                    {
                        yield return node;
                    }
                }
            }
        }
        
        private static IEnumerable<XElement> CellContent(XElement wordcell, Dictionary<string, XElement> styles)
        {
            XElement[] listelems = new XElement[10];
            XElement rootlist = null;
            int lastilvl = -1;

            foreach (XElement p in wordcell.Elements(xmlns.w + "p"))
            {
                if (p.Value != "")
                {
                    XElement paraprops = p.Element(xmlns.w + "pPr");
                    List<XElement> pstyles = new List<XElement>();

                    if (paraprops != null && 
                        paraprops.Elements(xmlns.w + "pStyle").Select(ps => ps.Attribute(xmlns.w + "val")).Any(ps => styles.ContainsKey(ps.Value)))
                    {
                        pstyles.Add(styles[paraprops.Element(xmlns.w + "pStyle").Attribute(xmlns.w + "val").Value]);
                    }

                    int numid = 0;
                    int ilvl = -1;

                    if (paraprops != null)
                    {
                        XElement[] numprs = paraprops.Elements(xmlns.w + "numPr").ToArray();

                        XElement numidelem = numprs.SelectMany(npr => npr.Elements(xmlns.w + "numId"))
                                                   .FirstOrDefault(nid => nid.Attribute(xmlns.w + "val") != null);

                        XElement ilvlelem = numprs.SelectMany(npr => npr.Elements(xmlns.w + "ilvl"))
                                                   .FirstOrDefault(ilv => ilv.Attribute(xmlns.w + "val") != null);

                        if (numidelem != null)
                        {
                            Int32.TryParse(numidelem.Attribute(xmlns.w + "val").Value, out numid);
                        }

                        if (ilvlelem != null)
                        {
                            Int32.TryParse(ilvlelem.Attribute(xmlns.w + "val").Value, out ilvl);
                        }
                    }

                    while (pstyles.Count != 0 && (numid == 0 || ilvl < 0))
                    {
                        List<XElement> _pstyles = pstyles;
                        pstyles = new List<XElement>();

                        foreach (XElement style in _pstyles)
                        {
                            List<XElement> numprs  = style.Elements(xmlns.w + "pPr")
                                                      .SelectMany(ppr => ppr.Elements(xmlns.w + "numPr"))
                                                      .ToList();

                            XElement numidelem = numprs.SelectMany(npr => npr.Elements(xmlns.w + "numId"))
                                                       .FirstOrDefault(nid => nid.Attribute(xmlns.w + "val") != null);

                            XElement ilvlelem  = numprs.SelectMany(npr => npr.Elements(xmlns.w + "ilvl"))
                                                       .FirstOrDefault(ilv => ilv.Attribute(xmlns.w + "val") != null);

                            if (numidelem != null && numid == 0)
                            {
                                Int32.TryParse(numidelem.Attribute(xmlns.w + "val").Value, out numid);
                            }

                            if (ilvlelem != null && ilvl < 0)
                            {
                                Int32.TryParse(ilvlelem.Attribute(xmlns.w + "val").Value, out ilvl);
                            }

                            foreach (XElement basedon in style.Elements(xmlns.w + "basedOn"))
                            {
                                string stylename = basedon.Attributes(xmlns.w + "val").Select(a => a.Value).FirstOrDefault();

                                if (styles.ContainsKey(stylename))
                                {
                                    pstyles.Add(styles[stylename]);
                                }
                            }
                        }
                    }

                    if (numid != 0)
                    {
                        if (ilvl < 0)
                        {
                            ilvl = 0;
                        }

                        if (rootlist == null)
                        {
                            rootlist = new XElement(ns_lasd + "ul");
                            listelems[0] = rootlist;
                            lastilvl = 0;
                        }

                        if (ilvl > lastilvl)
                        {
                            for (int i = lastilvl; i < ilvl; i++)
                            {
                                XElement listelem = new XElement(ns_lasd + "ul");
                                listelems[i].Add(listelem);
                                listelems[i + 1] = listelem;
                            }
                        }

                        listelems[ilvl].Add(new XElement(ns_lasd + "li", ParagraphContent(p, styles)));
                    }
                    else
                    {
                        if (rootlist != null)
                        {
                            yield return rootlist;
                            rootlist = null;
                            lastilvl = -1;
                        }

                        yield return new XElement(ns_lasd + "p", ParagraphContent(p, styles));
                    }
                }
            }

            if (rootlist != null)
            {
                yield return rootlist;
            }
        }

        private static WordTable GetTable(XElement tbl, Dictionary<string, XElement> styles)
        {
            WordTable table = new WordTable();
            XElement[] tblrows = tbl.Elements(xmlns.w + "tr").ToArray();
            int nrcols = tbl.Element(xmlns.w + "tblGrid").Elements(xmlns.w + "gridCol").Count();
            int nrrows = tblrows.Length;
            table.Cells = new WordTableCell[nrrows, nrcols];
            WordTableCell[] columns = new WordTableCell[nrcols];
            int trno = 0;

            foreach (XElement tblrow in tblrows)
            {
                int tcno = 0;

                foreach (XElement tblcell in tblrow.Elements(xmlns.w + "tc"))
                {
                    XElement cellprops = tblcell.Element(xmlns.w + "tcPr");
                    XElement vMerge = cellprops.Element(xmlns.w + "vMerge");
                    bool dovMerge = false;

                    if (vMerge != null && !vMerge.Attributes(xmlns.w + "val").Any(v => v.Value == "restart"))
                    {
                        dovMerge = true;
                    }

                    XElement gridSpan = cellprops.Element(xmlns.w + "gridSpan");
                    int colspan = 1;

                    if (gridSpan != null && gridSpan.Attribute(xmlns.w + "val") != null)
                    {
                        Int32.TryParse(gridSpan.Attribute(xmlns.w + "val").Value, out colspan);
                    }

                    if (dovMerge)
                    {
                        columns[tcno].RowSpan++;
                        table.Cells[trno, tcno] = columns[tcno];
                    }
                    else
                    {
                        XElement[] paragraphs = CellContent(tblcell, styles).ToArray();
                        WordTableCell cell = new WordTableCell
                        {
                            Row = trno,
                            Col = tcno,
                            RowSpan = 1,
                            ColSpan = colspan,
                            Paragraphs = paragraphs,
                        };

                        columns[tcno] = cell;
                        table.Cells[trno, tcno] = cell;
                    }

                    tcno += colspan;
                }
                trno++;
            }

            return table;
        }

        private static KeyLearningArea ProcessKLA(string yearLevel, string yearLevelId, string subject, string subjectId, WordTable[] tables)
        {
            KeyLearningArea kla = new KeyLearningArea
            {
                YearLevel = yearLevel,
                YearLevelID = yearLevelId,
                Subject = subject,
                SubjectID = subjectId,
                Groups = new List<AchievementRowGroup>(),
                Terms = new List<TermDefinition>()
            };

            foreach (WordTable table in tables)
            {
                int cols = table.Columns;
                int rows = table.Rows;

                if (cols >= 7)
                {
                    //kla.AchievementLevels = Enumerable.Range(0, cols).Select(c => table.Cells[0, c]).Where(c => c != null).Reverse().Take(5).Reverse().Select(c => c.Text).ToList();

                    for (int r = 1; r < rows; r++)
                    {
                        List<WordTableCell> cells = Enumerable.Range(0, cols).Select(c => table.Cells[r, c]).Where(c => c != null).Reverse().ToList();
                        List<string> groups = cells.Skip(5).Reverse().Select(c => c.Text.Replace("\n", " ")).ToList();
                        WordTableCell[] descs = cells.Take(5).Reverse().ToArray();

                        if (descs.Length == 5 && descs.Aggregate((int?)null, (a,d) => (a == null || d.RowSpan == a) ? d.RowSpan : -1) >= 1)
                        {
                            AchievementRowGroup grp = new AchievementRowGroup
                            {
                                Name = kla.YearLevel + " " + kla.Subject,
                                Groups = kla.Groups,
                                Rows = null,
                                Id = kla.YearLevelID + "::" + kla.SubjectID + "::"
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
                }
                else if (cols == 2 && table.Cells[0, 0].Text == "Term")
                {
                    for (int r = 1; r < rows; r++)
                    {
                        List<string> keywords = table.Cells[r, 0].Text.Split(',', ';').Select(k => k.Trim().Replace('\xA0', ' ')).ToList();
                        string name = (keywords.FirstOrDefault(k => k.EndsWith("*")) ?? keywords.First()).ToLower();
                        XElement[] elements = table.Cells[r, 1].Paragraphs;

                        if (name == "")
                        {
                            TermDefinition term = kla.Terms[kla.Terms.Count - 1];
                            term.Description.Elements = term.Description.Elements.Concat(elements).ToArray();
                        }
                        else
                        {
                            kla.Terms.Add(new TermDefinition
                            {
                                Name = name.TrimEnd('*'),
                                Keywords = keywords.Select(k => k.TrimEnd('*')).ToList(),
                                Description = new FormattedText { Elements = table.Cells[r, 1].Paragraphs }
                            });
                        }
                    }
                }
            }

            kla.FindTerms();

            return kla;
        }

        private static string GetDocVersion(Package pkg)
        {
            PackageFile corepropspart = pkg.GetRelation(PackageRelation.coreProperties);
            XElement coreprops = corepropspart.XmlDocument.Root;
            return coreprops.Elements(xmlns.dcterms + "modified").Select(e => e.Value).SingleOrDefault();
        }

        private static KeyLearningArea ProcessKLA(string yearLevel, string yearLevelId, string subject, string subjectId, Package pkg)
        {
            PackageFile docpart = pkg.GetRelation(PackageRelation.officeDocument);
            PackageFile stylespart = docpart.GetRelation(PackageRelation.styles);
            XElement body = docpart.XmlDocument.Root.Element(xmlns.w + "body");
            Dictionary<string, XElement> styles = stylespart.XmlDocument.Root.Elements(xmlns.w + "style").ToDictionary(s => s.Attribute(xmlns.w + "styleId").Value, s => s);
            XElement[] tables = body.Elements(xmlns.w + "tbl").ToArray();
            WordTable[] wordtables = tables.Select(t => GetTable(t, styles)).ToArray();
            KeyLearningArea kla = ProcessKLA(yearLevel, yearLevelId, subject, subjectId, wordtables);
            kla.Version = GetDocVersion(pkg);
            return kla;
        }

        private static void Main(string[] args)
        {
            GradeList gradelist = new GradeList { Grades = new List<Grade>() };
            Dictionary<string, Grade> grades = new Dictionary<string, Grade>();
            WebClient webclient = new WebClient();
            ProcessConfig();

            foreach (KeyValuePair<string, string> grade_kvp in YearLevels)
            {
                Grade grade = new Grade
                {
                    YearLevelID = grade_kvp.Key,
                    YearLevel = grade_kvp.Value,
                    Levels = AchievementLevels[GradeAchievementLevelRefs[grade_kvp.Key]].Select(kvp => new AchievementLevel { Abbreviation = kvp.Key, Name = kvp.Value }).ToList(),
                    KLAs = new List<KeyLearningAreaReference>()
                };

                grades[grade_kvp.Key] = grade;

                gradelist.Grades.Add(grade);
            }

            foreach (KeyValuePair<string, string[]> klagrp_kvp in SubjectGroupings)
            {
                string klagroupname = klagrp_kvp.Key;
                Dictionary<string, string[]> years = YearLevelGroupings[klagroupname];

                foreach (string klaname in klagrp_kvp.Value)
                {
                    foreach (KeyValuePair<string, string[]> gradegrp_kvp in years)
                    {
                        string gradegrpname = gradegrp_kvp.Key;
                        string filename = "ac_" + klaname + "_" + gradegrpname + "_se.docx";
                        string filepath = Path.Combine(Environment.CurrentDirectory, filename);

                        if (!File.Exists(filepath))
                        {
                            webclient.DownloadFile("https://www.qcaa.qld.edu.au/downloads/p_10/" + filename, filepath);
                        }

                        if (File.Exists(filepath))
                        {
                            foreach (string gradename in gradegrp_kvp.Value)
                            {
                                Console.WriteLine("Processing {0} {1} ({2})", YearLevels[gradename], Subjects[klaname], filename);

                                Package pkg = Package.Load(filepath);
                                KeyLearningArea kla = ProcessKLA(YearLevels[gradename], gradename, Subjects[klaname], klaname, pkg);
                                string xmlname = String.Format("{0}-{1}.xml", kla.YearLevelID, kla.SubjectID);

                                grades[gradename].KLAs.Add(new KeyLearningAreaReference
                                {
                                    SubjectID = kla.SubjectID,
                                    Subject = kla.Subject,
                                    Filename = xmlname,
                                    Version = kla.Version,
                                    Hash = kla.GetHash()
                                });

                                kla.ToXDocument().Save(xmlname);
                            }
                        }
                    }
                }
            }

            gradelist.ToXDocument().Save("grades.xml");
        }
    }
}
