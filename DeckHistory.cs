using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

sealed class DeckHistory {
    public sealed class Pair {public string Presenter,Audience;}
    readonly List<Pair> pairs=new List<Pair>();readonly string file;
    public DeckHistory(string file){this.file=file;try{var doc=new XmlDocument{XmlResolver=null};doc.Load(file);foreach(XmlNode node in doc.SelectNodes("/history/pair")){string a=node.Attributes["presenter"].Value,b=node.Attributes["audience"].Value;if(!String.IsNullOrWhiteSpace(a)&&!String.IsNullOrWhiteSpace(b))pairs.Add(new Pair{Presenter=a,Audience=b});if(pairs.Count>=20)break;}}catch{}}
    static bool Same(string a,string b){return String.Equals(a,b,StringComparison.OrdinalIgnoreCase);}
    public string Partner(string path){foreach(var pair in pairs){if(Same(path,pair.Presenter))return pair.Audience;if(Same(path,pair.Audience))return pair.Presenter;}return null;}
    public string[] Files(){var files=new List<string>();foreach(var pair in pairs)foreach(var path in new[]{pair.Presenter,pair.Audience})if(!files.Exists(p=>Same(p,path)))files.Add(path);return files.ToArray();}
    public void Remember(string presenter,string audience){presenter=Path.GetFullPath(presenter);audience=Path.GetFullPath(audience);pairs.RemoveAll(p=>Same(p.Presenter,presenter)&&Same(p.Audience,audience)||Same(p.Presenter,audience)&&Same(p.Audience,presenter));pairs.Insert(0,new Pair{Presenter=presenter,Audience=audience});if(pairs.Count>20)pairs.RemoveRange(20,pairs.Count-20);Save();}
    void Save(){var doc=new XmlDocument();var root=doc.CreateElement("history");doc.AppendChild(root);foreach(var pair in pairs){var node=doc.CreateElement("pair");node.SetAttribute("presenter",pair.Presenter);node.SetAttribute("audience",pair.Audience);root.AppendChild(node);}string temp=file+".tmp";doc.Save(temp);if(File.Exists(file))File.Replace(temp,file,null);else File.Move(temp,file);}
}

static class Preferences {
    public static string Root{get{string path=AppDomain.CurrentDomain.GetData("DualSlide.SettingsDirectory") as string;if(String.IsNullOrEmpty(path))path=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"DualSlide");try{Directory.CreateDirectory(path);return path;}catch{return AppDomain.CurrentDomain.BaseDirectory;}}}
}
