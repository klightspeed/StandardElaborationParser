using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.IO;
using Google.Apis.Drive;
using Google.GData.Spreadsheets;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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

    class Program
    {
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

        static IEnumerable<XNode> ParagraphContent(Paragraph para, Dictionary<string, Style> styles)
        {
            foreach (Run run in para.Elements<Run>())
            {
                foreach (Text text in run.Elements<Text>())
                {
                    yield return new XText(text.Text);
                }
            }
        }
        
        static IEnumerable<XElement> CellContent(TableCell wordcell, Dictionary<string, Style> styles)
        {
            XElement list = null;

            foreach (Paragraph p in wordcell.Elements<Paragraph>())
            {
                if (p.InnerText != "")
                {
                    ParagraphProperties paraprops = p.ParagraphProperties;
                    Style style = null;

                    if (paraprops != null && 
                        paraprops.ParagraphStyleId != null && 
                        paraprops.ParagraphStyleId.Val != null && 
                        styles.ContainsKey(paraprops.ParagraphStyleId.Val))
                    {
                        style = styles[paraprops.ParagraphStyleId.Val];
                    }

                    int numid = 0;

                    if (style != null && 
                        style.StyleParagraphProperties != null && 
                        style.StyleParagraphProperties.NumberingProperties != null && 
                        style.StyleParagraphProperties.NumberingProperties.NumberingId != null &&
                        style.StyleParagraphProperties.NumberingProperties.NumberingId.Val != null)
                    {
                        numid = style.StyleParagraphProperties.NumberingProperties.NumberingId.Val.Value;
                    }

                    if (paraprops != null &&
                        paraprops.NumberingProperties != null &&
                        paraprops.NumberingProperties.NumberingId != null &&
                        paraprops.NumberingProperties.NumberingId.Val != null)
                    {
                        numid = paraprops.NumberingProperties.NumberingId.Val.Value;
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

        static string CellText(XElement[] paragraphs)
        {
            List<string> lines = new List<string>();

            foreach (XElement para in paragraphs)
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

        static WordTable GetTable(Table tbl, Dictionary<string, Style> styles)
        {
            WordTable table = new WordTable();
            TableRow[] tblrows = tbl.Elements<TableRow>().ToArray();
            int nrcols = tbl.Elements<TableGrid>().Single().Elements<GridColumn>().Count();
            int nrrows = tblrows.Length;
            table.Cells = new WordTableCell[nrrows, nrcols];
            WordTableCell[] columns = new WordTableCell[nrcols];
            int trno = 0;

            foreach (TableRow tblrow in tblrows)
            {
                int tcno = 0;

                foreach (TableCell tblcell in tblrow.Elements<TableCell>())
                {
                    TableCellProperties cellprops = tblcell.Elements<TableCellProperties>().Single();
                    VerticalMerge vMerge = cellprops.Elements<VerticalMerge>().SingleOrDefault();
                    bool dovMerge = false;

                    if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
                    {
                        dovMerge = true;
                    }

                    GridSpan gridSpan = cellprops.Elements<GridSpan>().FirstOrDefault();
                    int colspan = 1;

                    if (gridSpan != null && gridSpan.Val != null)
                    {
                        colspan = gridSpan.Val.Value;
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
                            Text = CellText(paragraphs)
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

        static XElement FindTerms(XElement element, Dictionary<string, string> terms)
        {
            XElement ret = new XElement(element.Name, element.Attributes());

            foreach (XNode node in element.Nodes())
            {
                if (node is XElement)
                {
                    ret.Add(FindTerms((XElement)node, terms));
                }
                else if (node is XText)
                {
                    string text = ((XText)node).Value;
                    string lowertext = text.ToLower();
                    int startpos = 0;

                    do
                    {
                        int matchpos = text.Length;
                        int matchlen = 0;
                        string matchname = null;

                        foreach (KeyValuePair<string, string> term in terms)
                        {
                            int tmatchpos = lowertext.IndexOf(term.Key, startpos);

                            if (tmatchpos >= startpos && 
                                (tmatchpos < matchpos ||
                                 (tmatchpos == matchpos && term.Key.Length > matchlen)))
                            {
                                matchpos = tmatchpos;
                                matchlen = term.Key.Length;
                                matchname = term.Value;
                            }
                        }

                        if (matchlen != 0)
                        {
                            if (matchpos != startpos)
                            {
                                ret.Add(new XText(text.Substring(startpos, matchpos - startpos)));
                            }

                            ret.Add(new XElement(ns_lasd + "term",
                                new XAttribute("name", matchname),
                                new XText(text.Substring(matchpos, matchlen))
                            ));
                        }
                        else
                        {
                            ret.Add(new XText(text.Substring(startpos, text.Length - startpos)));
                        }

                        startpos = matchpos + matchlen;
                    }
                    while (startpos < text.Length);
                }
                else
                {
                    ret.Add(node);
                }
            }

            return ret;
        }

        static void FindTerms(AchievementRowGroup group, Dictionary<string, string> terms)
        {
            foreach (AchievementRowGroup grp in group.Groups)
            {
                FindTerms(grp, terms);
            }

            foreach (AchievementRow row in group.Rows)
            {
                foreach (FormattedText text in row.Descriptors)
                {
                    text.Elements = text.Elements.Select(e => FindTerms(e, terms)).ToArray();
                }
            }
        }

        static void FindTerms(KeyLearningArea kla)
        {
            Dictionary<string, string> terms = kla.Terms.SelectMany(t => t.Keywords.Select(k => new { keyword = k.ToLower(), name = t.Name })).ToDictionary(kn => kn.keyword, kn => kn.name);

            foreach (AchievementRowGroup group in kla.Groups)
            {
                FindTerms(group, terms);
            }
        }

        static KeyLearningArea ProcessKLA(string yearLevel, string yearLevelId, string subject, string subjectId, WordTable[] tables)
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

                        if (descs.Length == 5 && descs.All(d => d.ColSpan == 1 && d.RowSpan == 1))
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

            FindTerms(kla);

            return kla;
        }

        static void Main(string[] args)
        {
            foreach (KeyValuePair<string, string> grade_kvp in YearLevels)
            {
                foreach (KeyValuePair<string, string> subject_kvp in Subjects)
                {
                    string filename = Path.Combine(Environment.CurrentDirectory, @"ac_" + subject_kvp.Key + "_" + grade_kvp.Key + "_se");
                    Console.WriteLine("Processing {0} {1} ({2})", grade_kvp.Value, subject_kvp.Value, filename);

                    WordprocessingDocument doc = WordprocessingDocument.Open(filename + ".docx", false);
                    Body body = doc.MainDocumentPart.Document.Body;
                    Dictionary<string, Style> styles = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().ToDictionary(s => s.StyleId.Value, s => s);
                    Table[] tables = body.Elements<Table>().ToArray();
                    WordTable[] wordtables = tables.Select(t => GetTable(t, styles)).ToArray();

                    KeyLearningArea kla = ProcessKLA(grade_kvp.Value, grade_kvp.Key, subject_kvp.Value, subject_kvp.Key, wordtables);
                    kla.ToXDocument().Save(String.Format("{0}-{1}.xml", kla.YearLevelID, kla.SubjectID));
                }
            }
        }
    }
}
