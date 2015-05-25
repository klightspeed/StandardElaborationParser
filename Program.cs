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
                        foreach (XElement li in para.Elements().Where(el => el.Name.LocalName == "li"))
                        {
                            lines.Add(" • " + li.Value);
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
        private static XNamespace ns_lasd = "http://tempuri.org/XmlLasdDatabase.xsd";

        private static Dictionary<string, string> YearLevels = new Dictionary<string, string>
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

        private static Dictionary<string, string> Subjects = new Dictionary<string, string>
        {
            { "eng", "English" },
            { "math", "Mathematics" },
            { "geog", "Geography" },
            { "hist", "History" },
            { "sci", "Science" },
            { "hpe", "Health and Physical Education" },
            { "enb", "Economics and Business" },
            { "arts_dance", "The Arts: Dance" },
            { "arts_drama", "The Arts: Drama" },
            { "arts_media", "The Arts: Media Arts" },
            { "arts_music", "The Arts: Music" },
            { "arts_visual", "The Arts: Visual Arts" }
        };

        private static Dictionary<string, Dictionary<string, string>> AchievementLevels = new Dictionary<string, Dictionary<string, string>>
        {
            { "p-2", new Dictionary<string, string> {
                { "AP", "Applying" },
                { "MC", "Making Connections" },
                { "WW", "Working With" },
                { "EX", "Exploring" },
                { "BA", "Becoming Aware" }
            } },
            { "3-10", new Dictionary<string, string> {
                { "A", "A" },
                { "B", "B" },
                { "C", "C" },
                { "D", "D" },
                { "E", "E" }
            } }
        };

        private static Dictionary<string, string> GradeAchievementLevelRefs = new Dictionary<string, string>
        {
            { "prep", "p-2" },
            { "yr1", "p-2" },
            { "yr2", "p-2" },
            { "yr3", "3-10" },
            { "yr4", "3-10" },
            { "yr5", "3-10" },
            { "yr6", "3-10" },
            { "yr7", "3-10" },
            { "yr8", "3-10" },
            { "yr9", "3-10" },
            { "yr10", "3-10" }
        };

        private static IEnumerable<XNode> ParagraphContent(XElement para, Dictionary<string, XElement> styles)
        {
            foreach (XElement run in para.Elements(xmlns.w + "r"))
            {
                foreach (XElement text in run.Elements(xmlns.w + "t"))
                {
                    yield return new XText(text.Value);
                }
            }
        }
        
        private static IEnumerable<XElement> CellContent(XElement wordcell, Dictionary<string, XElement> styles)
        {
            XElement list = null;

            foreach (XElement p in wordcell.Elements(xmlns.w + "p"))
            {
                if (p.Value != "")
                {
                    XElement paraprops = p.Element(xmlns.w + "pPr");
                    XElement style = null;

                    if (paraprops != null && 
                        paraprops.Elements(xmlns.w + "pStyle").Select(ps => ps.Attribute(xmlns.w + "val")).Any(ps => styles.ContainsKey(ps.Value)))
                    {
                        style = styles[paraprops.Element(xmlns.w + "pStyle").Attribute(xmlns.w + "val").Value];
                    }

                    int numid = 0;

                    if (style != null && 
                        style.Elements(xmlns.w + "pPr").SelectMany(ppr => ppr.Elements(xmlns.w + "numPr")).SelectMany(npr => npr.Elements(xmlns.w + "numId")).Any(nid => nid.Attribute(xmlns.w + "val") != null))
                    {
                        Int32.TryParse(style.Element(xmlns.w + "pPr").Element(xmlns.w + "numPr").Element(xmlns.w + "numId").Attribute(xmlns.w + "val").Value, out numid);
                    }

                    if (paraprops != null &&
                        paraprops.Elements(xmlns.w + "numPr").SelectMany(npr => npr.Elements(xmlns.w + "numId")).Any(nid => nid.Attribute(xmlns.w + "val") != null))
                    {
                        Int32.TryParse(paraprops.Element(xmlns.w + "numPr").Element(xmlns.w + "numId").Attribute(xmlns.w + "val").Value, out numid);
                    }

                    if (numid != 0)
                    {
                        if (list == null)
                        {
                            list = new XElement(ns_lasd + "ul");
                        }

                        list.Add(new XElement(ns_lasd + "li", ParagraphContent(p, styles)));
                    }
                    else
                    {
                        if (list != null)
                        {
                            yield return list;
                            list = null;
                        }

                        yield return new XElement(ns_lasd + "p", ParagraphContent(p, styles));
                    }
                }
            }

            if (list != null)
            {
                yield return list;
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

            foreach (KeyValuePair<string, string> grade_kvp in YearLevels)
            {
                Grade grade = new Grade
                {
                    YearLevelID = grade_kvp.Key,
                    YearLevel = grade_kvp.Value,
                    Levels = AchievementLevels[GradeAchievementLevelRefs[grade_kvp.Key]].Select(kvp => new AchievementLevel { Abbreviation = kvp.Key, Name = kvp.Value }).ToList(),
                    KLAs = new List<KeyLearningAreaReference>()
                };

                foreach (KeyValuePair<string, string> subject_kvp in Subjects)
                {
                    string filename = Path.Combine(Environment.CurrentDirectory, @"ac_" + subject_kvp.Key + "_" + grade_kvp.Key + "_se.docx");

                    if (File.Exists(filename))
                    {
                        Console.WriteLine("Processing {0} {1} ({2})", grade_kvp.Value, subject_kvp.Value, filename);

                        Package pkg = Package.Load(filename);
                        KeyLearningArea kla = ProcessKLA(grade_kvp.Value, grade_kvp.Key, subject_kvp.Value, subject_kvp.Key, pkg);
                        string xmlname = String.Format("{0}-{1}.xml", kla.YearLevelID, kla.SubjectID);

                        grade.KLAs.Add(new KeyLearningAreaReference
                        {
                            SubjectID = kla.SubjectID,
                            Subject = kla.Subject,
                            Filename = xmlname,
                            Version = kla.Version
                        });

                        kla.ToXDocument().Save(xmlname);
                    }
                }

                gradelist.Grades.Add(grade);
            }

            gradelist.ToXDocument().Save("grades.xml");
        }
    }
}
