using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TSVCEO.XmlLasdDatabase;
using TSVCEO.PDF;
using TSVCEO.PDF.Primitives;
using System.Drawing;
using System.Xml.Linq;

namespace StandardElaborationParser
{
    public class StandardElaborationPDFParser
    {
        protected class xmlns
        {
            public static readonly XNamespace lasd = "http://gtmj.tsv.catholic.edu.au/XmlLasdDatabase.xsd";
        }

        protected class PDFParagraph
        {
            public PDFContentBlock Content;
            public XElement Element;
        }

        protected class PDFListEntry : PDFParagraph
        {
            public List<PDFParagraph> Paragraphs;
        }

        protected class PDFList : PDFParagraph
        {
            public List<PDFListEntry> Entries;
        }

        protected class PDFTableCell
        {
            public PDFContentBlock Content;
            public List<PDFParagraph> Paragraphs;
            public List<XElement> Elements;
            public RectangleF CropBox { get { return Content == null ? RectangleF.Empty : Content.CropBox; } }
            public PointF TextPos { get { return Content == null ? PointF.Empty : Content.TextPos; } }
            public int ColSpan;
            public int RowSpan;
        }

        protected class PDFTableHeaderCell : PDFTableCell
        {
        }

        protected class PDFTableRow
        {
            public PDFContentBlock Artifact;
            public PDFContentBlock Content;
            public List<PDFTableCell> Cells;
            public float RowTop { get { return Cells.Max(c => Math.Max(c.CropBox.Y + c.CropBox.Height, c.TextPos.Y)); } }
        }

        protected class PDFTable
        {
            public PDFContentBlock Content;
            public List<PDFTableRow> Rows;
        }

        protected static IEnumerable<PDFParagraph> EnumerateListParagraphs(PDFContentBlock block)
        {
            if (block.BlockType.Name == "LBody")
            {
                yield return new PDFParagraph
                {
                    Content = block,
                };
            }
        }

        protected static IEnumerable<PDFListEntry> EnumerateListEntries(PDFContentBlock block)
        {
            if (block.BlockType.Name == "LI")
            {
                string text = block.Text.Trim();
                if (text.StartsWith("\x95"))
                {
                    text = text.Substring(1).Trim();
                }

                yield return new PDFListEntry
                {
                    Content = block,
                    Paragraphs = block.Content.OfType<PDFContentBlock>().SelectMany(p => EnumerateListParagraphs(p)).ToList(),
                    Element = new XElement(xmlns.lasd + "li", text)
                };
            }
        }

        protected static IEnumerable<PDFParagraph> EnumerateParagraphs(PDFContentBlock block)
        {
            if (block.BlockType.Name == "P")
            {
                if (block.Text.Trim() != "")
                {
                    yield return new PDFParagraph { Content = block, Element = new XElement(xmlns.lasd + "p", block.Text.Trim()) };
                }
            }
            else if (block.BlockType.Name == "L")
            {
                PDFList list = new PDFList
                {
                    Content = block,
                    Entries = block.Content.OfType<PDFContentBlock>().SelectMany(p => EnumerateListEntries(p)).ToList()
                };

                list.Element = new XElement(xmlns.lasd + "ul", list.Entries.Select(e => e.Element));

                yield return list;
            }
            else if (block.BlockType.Name != "Artifact")
            {
                foreach (PDFContentBlock blk in block.Content.OfType<PDFContentBlock>())
                {
                    foreach (PDFParagraph p in EnumerateParagraphs(blk))
                    {
                        yield return p;
                    }
                }
            }
        }

        protected static IEnumerable<PDFTableCell> EnumerateTableCells(PDFContentBlock block)
        {
            if (block.Content.Count != 0)
            {
                if (block.BlockType.Name == "TD")
                {
                    PDFTableCell cell = new PDFTableCell
                    {
                        Content = block,
                        Paragraphs = block.Content.OfType<PDFContentBlock>().SelectMany(p => EnumerateParagraphs(p)).ToList()
                    };

                    if (block.Attributes != null && block.Attributes.Dict != null)
                    {
                        PDFDictionary attrs = block.Attributes.Dict;
                        PDFInteger rowspan;
                        PDFInteger colspan;

                        if (attrs.TryGet("RowSpan", out rowspan))
                        {
                            cell.RowSpan = (int)rowspan.Value;
                        }

                        if (attrs.TryGet("ColSpan", out colspan))
                        {
                            cell.ColSpan = (int)colspan.Value;
                        }
                    }

                    cell.Elements = cell.Paragraphs.Select(p => p.Element).ToList();

                    yield return cell;
                }
                if (block.BlockType.Name == "TH")
                {
                    PDFTableHeaderCell cell = new PDFTableHeaderCell
                    {
                        Content = block,
                        Paragraphs = block.Content.OfType<PDFContentBlock>().SelectMany(p => EnumerateParagraphs(p)).ToList()
                    };

                    if (block.Attributes != null && block.Attributes.Dict != null)
                    {
                        PDFDictionary attrs = block.Attributes.Dict;
                        PDFInteger rowspan;
                        PDFInteger colspan;

                        if (attrs.TryGet("RowSpan", out rowspan))
                        {
                            cell.RowSpan = (int)rowspan.Value;
                        }

                        if (attrs.TryGet("ColSpan", out colspan))
                        {
                            cell.ColSpan = (int)colspan.Value;
                        }
                    }

                    cell.Elements = cell.Paragraphs.Select(p => p.Element).ToList();

                    yield return cell;
                }
                else if (block.BlockType.Name != "Artifact")
                {
                    foreach (PDFContentBlock blk in block.Content.OfType<PDFContentBlock>())
                    {
                        foreach (PDFTableCell cell in EnumerateTableCells(blk))
                        {
                            yield return cell;
                        }
                    }
                }
            }
        }

        protected static IEnumerable<PDFTableRow> EnumerateTableRows(PDFContentBlock block)
        {
            if (block.BlockType.Name == "TR")
            {
                PDFTableRow row = new PDFTableRow { Content = block, Cells = new List<PDFTableCell>() };

                foreach (PDFContentBlock blk in block.Content.OfType<PDFContentBlock>().Where(b => b.Content.Count != 0))
                {
                    if (blk.BlockType.Name == "Artifact")
                    {
                        if (row.Cells.Count != 0)
                        {
                            yield return row;
                        }

                        row = new PDFTableRow { Content = block, Cells = new List<PDFTableCell>() };
                        row.Artifact = blk;
                    }
                    else
                    {
                        row.Cells.AddRange(EnumerateTableCells(blk));
                    }
                }

                if (row.Cells.Count != 0)
                {
                    yield return row;
                }
            }
            else if (block.BlockType.Name != "Artifact")
            {
                foreach (PDFContentBlock blk in block.Content.OfType<PDFContentBlock>())
                {
                    foreach (PDFTableRow row in EnumerateTableRows(blk))
                    {
                        yield return row;
                    }
                }
            }
        }

        protected static IEnumerable<PDFTable> EnumerateTables(PDFContentBlock block)
        {
            if (block.BlockType.Name == "Table")
            {
                yield return new PDFTable {
                    Content = block,
                    Rows = block.Content.OfType<PDFContentBlock>().SelectMany(c => EnumerateTableRows(c)).ToList()
                };
            }
            else if (block.BlockType.Name != "Artifact")
            {
                foreach (PDFContentBlock blk in block.Content.OfType<PDFContentBlock>())
                {
                    foreach (PDFTable tbl in EnumerateTables(blk))
                    {
                        yield return tbl;
                    }
                }
            }
        }

        protected static KeyLearningArea ProcessKLA(string yearLevel, string yearLevelId, string subject, string subjectId, List<PDFTable> tables)
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

            foreach (PDFTable table in tables)
            {
                int cols = table.Rows.Max(r => r.Cells.Count);
                int rows = table.Rows.Count;

                if (cols >= 6)
                {
                    List<float> colx = new List<float>();
                    List<float> rowy = table.Rows.Select(r => r.RowTop).ToList();
                    int maxcol = 0;

                    // Find column widths
                    for (int rownum = 0; rownum < table.Rows.Count; rownum++)
                    {
                        PDFTableRow tr = table.Rows[rownum];
                        int colnum = 0;

                        foreach (PDFTableCell tc in tr.Cells)
                        {
                            float x = tc.CropBox.X;
                            float y = tc.CropBox.Y;

                            if (x != 0 && y != 0)
                            {
                                while (colnum < colx.Count && x > colx[colnum] + 20)
                                {
                                    colnum++;
                                }

                                if (colnum >= colx.Count)
                                {
                                    while (colnum >= colx.Count)
                                    {
                                        colx.Add(x);
                                    }
                                }
                                else if (x < colx[colnum] - 10)
                                {
                                    int c0 = colnum;

                                    while (c0 > 0 && x < colx[c0 - 1] - 10)
                                    {
                                        c0--;
                                    }

                                    for (; c0 <= colnum; c0++)
                                    {
                                        colx.Insert(c0, x);
                                    }
                                }
                                else if (x < colx[colnum])
                                {
                                    colx[colnum] = x;
                                }
                            }

                            colnum++;
                        }

                        if (colnum >= maxcol)
                        {
                            maxcol = colnum + 1;
                        }
                    }

                    PDFTableCell[][] cells = Enumerable.Range(0, rowy.Count).Select(i => new PDFTableCell[maxcol + 1]).ToArray();

                    // Find column and row spans
                    for (int rownum = 0; rownum < table.Rows.Count; rownum++)
                    {
                        PDFTableRow tr = table.Rows[rownum];
                        List<PDFTableCell> rcells = new List<PDFTableCell>();
                        int colnum = 0;

                        PDFTableCell ltc = null;

                        foreach (PDFTableCell tc in tr.Cells)
                        {
                            float x = tc.TextPos.X;
                            float y = tc.TextPos.Y;

                            if ((x != 0 && y != 0) || (tc.Content.Text != null && tc.Content.Text.Trim() != ""))
                            {
                                int cn = colnum;

                                if (x != 0)
                                {
                                    while (cn < colx.Count - 1 && x >= colx[cn + 1])
                                    {
                                        cn++;
                                    }
                                }

                                if (ltc != null && ltc.ColSpan == 0)
                                {
                                    ltc.ColSpan = Math.Max(cn + 1 - colnum, 1);
                                }

                                int rn = rownum;

                                if (y != 0)
                                {
                                    while (rn < rowy.Count - 1 && y < rowy[rn + 1])
                                    {
                                        rn++;
                                    }
                                }

                                if (tc.RowSpan == 0)
                                {
                                    tc.RowSpan = Math.Max(rn + 1 - rownum, 1);
                                }

                                for (int i = 0; i < tc.RowSpan; i++)
                                {
                                    cells[rownum + i][cn] = tc;
                                }

                                colnum = cn + 1;
                            }

                            ltc = tc;
                        }
                    }

                    List<string> lastgroups = new List<string>();

                    for (int rownum = 1; rownum < cells.Length; rownum++)
                    {
                        List<PDFTableCell> rcells = cells[rownum].Where(c => c != null).Reverse().ToList();
                        List<string> groups = rcells.Skip(5).Reverse().Select(c => c.Content.Text.Replace("\n", " ")).ToList();
                        PDFTableCell[] descs = rcells.Take(5).Reverse().ToArray();

                        if (groups.Count == 0 || (groups.Count < lastgroups.Count && groups.Select((g, i) => g == lastgroups[i]).All(g => g)))
                        {
                            groups = lastgroups;
                        }
                        else
                        {
                            lastgroups = groups;
                        }

                        if (descs.Length == 5 && !descs.Any(d => d is PDFTableHeaderCell) && descs.Aggregate((int?)null, (a, d) => (a == null || d.RowSpan == a) ? d.RowSpan : -1) >= 1)
                        {
                            AchievementRowGroup grp = new AchievementRowGroup
                            {
                                Name = kla.YearLevel + " " + kla.Subject,
                                Groups = kla.Groups,
                                Rows = new List<AchievementRow>(),
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
                                Descriptors = descs.Select(d => new FormattedText { Elements = d.Elements.ToArray() }).ToList(),
                                Id = grp.Id + (grp.Rows.Count + 1).ToString()
                            });
                        }
                    }

                }
                else if (cols == 2 && rows >= 2 && table.Rows[0].Cells[0].Content.Text.Trim() == "Term")
                {
                    for (int r = 1; r < rows; r++)
                    {
                        PDFTableRow row = table.Rows[r];
                        if (row.Cells.Count == 2 && row.Cells.All(c => c != null))
                        {
                            List<string> keywords = row.Cells[0].Content.Text.Split(',', ';').Select(k => k.Trim().Replace('\xA0', ' ')).ToList();
                            string name = (keywords.FirstOrDefault(k => k.EndsWith("*")) ?? keywords.First()).ToLower();
                            XElement[] elements = row.Cells[1].Elements.ToArray();

                            if (name == "")
                            {
                                if (kla.Terms.Count != 0)
                                {
                                    TermDefinition term = kla.Terms[kla.Terms.Count - 1];
                                    term.Description.Elements = term.Description.Elements.Concat(elements).ToArray();
                                }
                            }
                            else
                            {
                                kla.Terms.Add(new TermDefinition
                                {
                                    Name = name.TrimEnd('*'),
                                    Keywords = keywords.Select(k => k.TrimEnd('*')).ToList(),
                                    Description = new FormattedText { Elements = elements }
                                });
                            }
                        }
                    }
                }
            }

            return kla;
        }

        public static KeyLearningArea ProcessKLA(string yearLevel, string yearLevelId, string subject, string subjectId, PDFDocument doc)
        {
            PDFContentBlock root = doc.StructTree;
            if (root != null)
            {
                List<PDFTable> tables = EnumerateTables(root).ToList();
                KeyLearningArea kla = ProcessKLA(yearLevel, yearLevelId, subject, subjectId, tables);
                if (doc.Info != null && doc.Info.Dict != null)
                {
                    PDFDictionary info = doc.Info.Dict;
                    PDFString moddate;
                    if (info.TryGet("ModDate", out moddate))
                    {
                        kla.Version = moddate.Value;
                    }
                }
                return kla;
            }
            else
            {
                Console.WriteLine("Document has no structural tree!!!");
                return null;
            }
        }

        public static KeyLearningArea ProcessKLA(string yearLevel, string yearLevelId, string subject, string subjectId, string filepath)
        {
            PDFDocument doc = PDFDocument.ParseDocument(filepath);
            return ProcessKLA(yearLevel, yearLevelId, subject, subjectId, doc);
        }
    }
}
