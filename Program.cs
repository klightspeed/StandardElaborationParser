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
    public class Program
    {
        private static Dictionary<string, string> YearLevels;

        private static Dictionary<string, string> Subjects;

        private static Dictionary<string, string> SubjectFormats;

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
            SubjectFormats =
                root.Element("Subjects")
                    .Elements("Subject")
                    .ToDictionary(
                        e => e.Attribute("id").Value,
                        e => e.Attribute("format").Value
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


        private static void Main(string[] args)
        {
            GradeList gradelist = new GradeList { Grades = new List<Grade>() };
            Dictionary<string, Grade> grades = new Dictionary<string, Grade>();
            WebClient webclient = new WebClient();
            Dictionary<string, Dictionary<string, KeyLearningAreaReference>> klarefs = new Dictionary<string, Dictionary<string, KeyLearningAreaReference>>();
            ProcessConfig();

            foreach (KeyValuePair<string, string[]> klagrp_kvp in SubjectGroupings)
            {
                string klagroupname = klagrp_kvp.Key;
                Dictionary<string, string[]> years = YearLevelGroupings[klagroupname];

                foreach (string klaname in klagrp_kvp.Value)
                {
                    foreach (KeyValuePair<string, string[]> gradegrp_kvp in years)
                    {
                        string format = SubjectFormats.ContainsKey(klaname) ? SubjectFormats[klaname] : "docx";
                        string gradegrpname = gradegrp_kvp.Key;
                        string filename = "ac_" + klaname + "_" + gradegrpname + "_se." + format;
                        string filepath = Path.Combine(Environment.CurrentDirectory, filename);
                        string sourceurl = "https://www.qcaa.qld.edu.au/downloads/p_10/" + filename;

                        if (!File.Exists(filepath))
                        {
                            webclient.DownloadFile(sourceurl, filepath);
                        }

                        if (File.Exists(filepath))
                        {
                            foreach (string gradename in gradegrp_kvp.Value)
                            {
                                Console.WriteLine("Processing {0} {1} ({2})", YearLevels[gradename], Subjects[klaname], filename);

                                KeyLearningArea kla = null;

                                if (format == "docx")
                                {
                                    kla = StandardElaborationDocxParser.ProcessKLA(YearLevels[gradename], gradename, Subjects[klaname], klaname, filepath);
                                }
                                else if (format == "pdf")
                                {
                                    kla = StandardElaborationPDFParser.ProcessKLA(YearLevels[gradename], gradename, Subjects[klaname], klaname, filepath);
                                }

                                if (kla != null)
                                {
                                    kla.SourceDocumentURL = sourceurl;
                                    string xmlname = String.Format("{0}-{1}.xml", kla.YearLevelID, kla.SubjectID);

                                    if (!klarefs.ContainsKey(gradename))
                                    {
                                        klarefs[gradename] = new Dictionary<string, KeyLearningAreaReference>();
                                    }

                                    klarefs[gradename][klaname] = new KeyLearningAreaReference
                                    {
                                        SubjectID = kla.SubjectID,
                                        Subject = kla.Subject,
                                        Filename = xmlname,
                                        Version = kla.Version,
                                        Hash = kla.GetHash()
                                    };

                                    kla.ToXDocument().Save(xmlname);
                                }
                            }
                        }
                    }
                }
            }

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

                if (klarefs.ContainsKey(grade_kvp.Key))
                {
                    foreach (KeyValuePair<string, string> kla_kvp in Subjects)
                    {
                        if (klarefs[grade_kvp.Key].ContainsKey(kla_kvp.Key))
                        {
                            grades[grade_kvp.Key].KLAs.Add(klarefs[grade_kvp.Key][kla_kvp.Key]);
                        }
                    }
                }
            }

            gradelist.ToXDocument().Save("grades.xml");
        }
    }
}
