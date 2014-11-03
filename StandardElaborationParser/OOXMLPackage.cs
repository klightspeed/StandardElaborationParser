using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Ionic.Zip;

namespace StandardElaborationParser
{
    public static class xmlns
    {
        public static readonly XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        public static readonly XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        public static readonly XNamespace bibliography = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
        public static readonly XNamespace extendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        public static readonly XNamespace dwg = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static readonly XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        public static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";
        public static readonly XNamespace dcmitype = "http://purl.org/dc/dcmitype/";
        public static readonly XNamespace dcterms = "http://purl.org/dc/terms/";
        public static readonly XNamespace ds = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
        public static readonly XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        public static readonly XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public static readonly XNamespace o = "urn:schemas-microsoft-com:office:office";
        public static readonly XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace sl = "http://schemas.openxmlformats.org/schemaLibrary/2006/main";
        public static readonly XNamespace thm15 = "http://schemas.microsoft.com/office/thememl/2012/main";
        public static readonly XNamespace v = "urn:schemas-microsoft-com:vml";
        public static readonly XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        public static readonly XNamespace w10 = "urn:schemas-microsoft-com:office:word";
        public static readonly XNamespace w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        public static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace wne = "http://schemas.microsoft.com/office/word/2006/wordml";
        public static readonly XNamespace wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
        public static readonly XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XNamespace wpc = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
        public static readonly XNamespace wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
        public static readonly XNamespace wpi = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
        public static readonly XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        public static readonly XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
    }

    public abstract class PackageEntry
    {
        public string Name { get; set; }
        public string ContentType { get; set; }
        public Package Package { get; set; }
        public PackageDirectory Parent { get; set; }
        public PackageRelation[] Relations { get; set; }

        public string Path
        {
            get
            {
                return (Parent == Package ? "" : Parent.Path) + "/" + Name;
            }
        }

        public abstract void LoadEntry(ZipEntry entry, string[] pathcomponents);
        public abstract void Save(ZipFile zip, string path);

        public void LoadRelations(PackageFile relationsFile)
        {
            XDocument doc = relationsFile.XmlDocument;
            List<PackageRelation> relations = new List<PackageRelation>();

            foreach (XElement relation in doc.Root.Elements(xmlns.relationships + "Relationship"))
            {
                string id = relation.Attribute("Id").Value;
                string type = relation.Attribute("Type").Value;
                string targetname = relation.Attribute("Target").Value;
                string targetmode = relation.Attributes("TargetMode").Select(a => a.Value).SingleOrDefault();
                PackageFile target = null;

                if (targetmode != "External")
                {
                    target = this.Parent[targetname] as PackageFile;
                }

                relations.Add(new PackageRelation
                {
                    Id = id,
                    Type = type,
                    TargetName = targetname,
                    Target = target,
                    IsExternal = targetmode == "External"
                });
            }

            this.Relations = relations.ToArray();
        }

        public IEnumerable<PackageFile> GetRelations(string type)
        {
            return Relations.Where(r => r.Type == type).Select(r => r.Target);
        }

        public PackageFile GetRelation(string type)
        {
            return GetRelations(type).Single();
        }
    }

    public class PackageDirectory : PackageEntry, IEnumerable<PackageEntry>
    {
        protected Dictionary<string, PackageEntry> Entries { get; set; }

        public PackageDirectory()
        {
            Entries = new Dictionary<string, PackageEntry>();
        }

        public override void LoadEntry(ZipEntry entry, string[] pathcomponents)
        {
            Name = pathcomponents[0];

            if (Entries.ContainsKey(pathcomponents[1].ToLower()))
            {
                Entries[pathcomponents[1].ToLower()].LoadEntry(entry, pathcomponents.Skip(1).ToArray());
            }
            else
            {
                if (pathcomponents.Length == 2)
                {
                    Entries[pathcomponents[1].ToLower()] = new PackageFile { Package = this.Package, Parent = this };
                    Entries[pathcomponents[1].ToLower()].LoadEntry(entry, pathcomponents.Skip(1).ToArray());
                }
                else
                {
                    Entries[pathcomponents[1].ToLower()] = new PackageDirectory { Name = pathcomponents[1], Package = this.Package, Parent = this };
                    Entries[pathcomponents[1].ToLower()].LoadEntry(entry, pathcomponents.Skip(1).ToArray());
                }
            }
        }

        public override void Save(ZipFile zip, string path)
        {
            path = (String.IsNullOrEmpty(path) ? "" : path + "/") + this.Name;

            foreach (PackageEntry entry in Entries.Values)
            {
                entry.Save(zip, path);
            }
        }

        protected PackageEntry GetEntry(string[] pathcomponents)
        {
            if (pathcomponents[0] == "..")
            {
                return Parent.GetEntry(pathcomponents.Skip(1).ToArray());
            }

            PackageEntry entry = Entries[pathcomponents[0].ToLower()];

            if (pathcomponents.Length == 1)
            {
                return entry;
            }
            else if (entry is PackageDirectory)
            {
                return ((PackageDirectory)entry).GetEntry(pathcomponents.Skip(1).ToArray());
            }
            else
            {
                throw new KeyNotFoundException();
            }
        }

        protected bool ContainsEntry(string[] pathcomponents)
        {
            if (Entries.ContainsKey(pathcomponents[0].ToLower()))
            {
                if (pathcomponents.Length == 1)
                {
                    return true;
                }
                else if (Entries[pathcomponents[0].ToLower()] is PackageDirectory)
                {
                    return ((PackageDirectory)Entries[pathcomponents[0].ToLower()]).ContainsEntry(pathcomponents.Skip(1).ToArray());
                }
            }

            return false;
        }

        public PackageEntry this[string name]
        {
            get
            {
                string[] pathcomponents = name.Split('\\', '/').Where(s => s != "").ToArray();

                if (name.StartsWith("/") || name.StartsWith("\\"))
                {
                    return Package.GetEntry(pathcomponents);
                }
                else
                {
                    return GetEntry(pathcomponents);
                }
            }
        }

        public IEnumerable<PackageFile> GetAllFiles(string extension)
        {
            foreach (PackageEntry entry in this.Entries.Values)
            {
                if (entry is PackageDirectory)
                {
                    foreach (PackageFile pkgfile in ((PackageDirectory)entry).GetAllFiles(extension))
                    {
                        yield return pkgfile;
                    }
                }
                else if (entry is PackageFile)
                {
                    if (entry.Name.EndsWith(extension))
                    {
                        yield return (PackageFile)entry;
                    }
                }
            }
        }

        public bool ContainsEntry(string name)
        {
            string[] pathcomponents = name.Split('\\', '/').Where(s => s != "").ToArray();
            return ContainsEntry(pathcomponents);
        }

        public IEnumerator<PackageEntry> GetEnumerator()
        {
            foreach (PackageEntry entry in Entries.Values)
            {
                yield return entry;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void LoadRelations()
        {
            foreach (PackageEntry pkgent in this)
            {
                if (this.ContainsEntry("_rels/.rels"))
                {
                    LoadRelations(this["_rels/.rels"] as PackageFile);
                }

                if (pkgent is PackageFile)
                {
                    PackageFile pkgfile = (PackageFile)pkgent;
                    if (this.ContainsEntry("_rels/" + pkgfile.Name + ".rels"))
                    {
                        pkgfile.LoadRelations(this["_rels/" + pkgfile.Name + ".rels"] as PackageFile);
                    }
                }
                else if (pkgent is PackageDirectory)
                {
                    ((PackageDirectory)pkgent).LoadRelations();
                }
            }
        }
    }

    public class PackageFile : PackageEntry
    {
        protected XDocument _XmlDocument;
        protected byte[] _Data;

        public virtual byte[] Data
        {
            get
            {
                if (_Data != null)
                {
                    return _Data;
                }
                else
                {
                    using (MemoryStream stream = new MemoryStream())
                    {
                        using (XmlTextWriter writer = new XmlTextWriter(stream, new UTF8Encoding(false)))
                        {
                            writer.Formatting = Formatting.None;
                            this.XmlDocument.Save(writer);
                        }
                        return stream.ToArray();
                    }
                }
            }
            set
            {
                _XmlDocument = null;
                _Data = value;
            }
        }

        public virtual XDocument XmlDocument
        {
            get
            {
                if (_XmlDocument != null)
                {
                    return _XmlDocument;
                }
                else if (_Data != null)
                {
                    _XmlDocument = XDocument.Load(new MemoryStream(_Data));
                    _Data = null;
                    return _XmlDocument;
                }
                else
                {
                    throw new InvalidOperationException();
                }
            }
        }

        public override void LoadEntry(ZipEntry entry, string[] pathcomponents)
        {
            Name = pathcomponents[0];

            using (MemoryStream memstrm = new MemoryStream())
            {
                entry.Extract(memstrm);
                Data = memstrm.ToArray();
            }
        }

        public override void Save(ZipFile zip, string path)
        {
            string name = (String.IsNullOrEmpty(path) ? "" : path + "/") + this.Name;

            zip.AddEntry(name, Data);
        }
    }

    public class PackageContentTypes : PackageFile
    {
        public override byte[] Data
        {
            get
            {
                return base.Data;
            }
            set
            {
                throw new InvalidOperationException();
            }
        }

        public override XDocument XmlDocument
        {
            get
            {
                return new XDocument(
                    new XElement(xmlns.contentTypes + "Types",
                        new XAttribute("xmlns", xmlns.contentTypes.NamespaceName),
                        new XElement(xmlns.contentTypes + "Default",
                            new XAttribute("Extension", "rels"),
                            new XAttribute("ContentType", PackageContentTypes.relationships)
                        ),
                        new XElement(xmlns.contentTypes + "Default",
                            new XAttribute("Extension", "xml"),
                            new XAttribute("ContentType", "application/xml")
                        ),
                        Package.GetAllFiles("xml").Select(f =>
                            f.ContentType == null ? null : new XElement(xmlns.contentTypes + "Override",
                                new XAttribute("PartName", f.Path),
                                new XAttribute("ContentType", f.ContentType)
                            )
                        )
                    )
                );
            }
        }

        public static string stylesWithEffects = "application/vnd.ms-word.stylesWithEffects+xml";
        public static string customXmlProperties = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
        public static string extendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        public static string theme = "application/vnd.openxmlformats-officedocument.theme+xml";
        public static string document = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        public static string endnotes = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
        public static string fontTable = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
        public static string footer = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
        public static string footnotes = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
        public static string header = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
        public static string numbering = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
        public static string settings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
        public static string styles = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
        public static string template = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
        public static string webSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
        public static string coreProperties = "application/vnd.openxmlformats-package.core-properties+xml";
        public static string relationships = "application/vnd.openxmlformats-package.relationships+xml";
        public static string xml = "application/xml";
    }

    public class Package : PackageDirectory
    {
        protected Dictionary<string, string> DefaultContentTypes { get; set; }

        protected void LoadContentTypes()
        {
            DefaultContentTypes = new Dictionary<string, string>();
            PackageFile file = this["[Content_Types].xml"] as PackageFile;
            XDocument doc = file.XmlDocument;

            foreach (XElement el in doc.Root.Elements(xmlns.contentTypes + "Default"))
            {
                string extension = el.Attribute("Extension").Value;
                string type = el.Attribute("ContentType").Value;
                DefaultContentTypes[extension] = type;
            }

            foreach (XElement el in doc.Root.Elements(xmlns.contentTypes + "Override"))
            {
                string name = el.Attribute("PartName").Value;
                string type = el.Attribute("ContentType").Value;

                this[name].ContentType = type;
            }

            Entries["[content_types].xml"] = new PackageContentTypes
            {
                Package = this,
                Parent = this,
                Name = "[Content_Types].xml",
                ContentType = null
            };
        }

        public static Package Load(string filename)
        {
            Package pkg = new Package();
            pkg.Package = pkg;
            pkg.Parent = pkg;

            using (ZipFile zip = new ZipFile(filename))
            {
                foreach (var entry in zip.Entries)
                {
                    string[] pathcomponents = new string[] { "" }.Concat(entry.FileName.Split('\\', '/').Where(s => s != "")).ToArray();
                    pkg.LoadEntry(entry, pathcomponents);
                }
            }

            pkg.LoadContentTypes();
            pkg.LoadRelations();

            return pkg;
        }

        public void SaveAs(string filename)
        {
            using (ZipFile zip = new ZipFile())
            {
                this.Save(zip, null);

                zip.Save(filename);
            }
        }
    }

    public class PackageRelation
    {
        public string Id { get; set; }
        public string Type { get; set; }
        public string TargetName { get; set; }
        public PackageFile Target { get; set; }
        public bool IsExternal { get; set; }

        public static string stylesWithEffects = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
        public static string customXml = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        public static string customXmlProps = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
        public static string endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
        public static string extendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        public static string fontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
        public static string footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
        public static string footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
        public static string header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        public static string numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        public static string officeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public static string settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
        public static string styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        public static string theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        public static string webSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
        public static string coreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        public static string hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    }
}
