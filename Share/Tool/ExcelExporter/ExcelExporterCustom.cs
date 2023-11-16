using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Security.Cryptography;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Bson.Serialization.Attributes;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ET.ExcelTool
{
    public enum ConfigType
    {
        c = 0,
        s = 1,
        cs = 2,
    }

    class HeadInfo
    {
        public string FieldDesc;
        public string FieldName;
        public string FieldType;
        public int FieldIndex;
        public Dictionary<string, string> FieldConfigs;

        public string FieldCS
        {
            get
            {
                if (FieldConfigs.TryGetValue("cs", out string cs))
                {
                    return cs;
                }
                return "cs";
            }
        }

        public HeadInfo(string desc, string name, string type, int index, Dictionary<string, string> fieldConfigs)
        {
            this.FieldDesc = desc;
            this.FieldName = name;
            this.FieldType = type;
            this.FieldIndex = index;
            this.FieldConfigs = fieldConfigs;
            FieldType = FieldType.Trim();
            if (FieldType.StartsWith("<") && FieldType.EndsWith(">"))
            {
                if (FieldType.Contains(","))
                {
                    FieldType = $"Dictionary{FieldType}";
                }
                else
                {
                    FieldType = $"{FieldType.TrimStart('<').TrimEnd('>')}[]";
                }
            }
        }
    }

    // 这里加个标签是为了防止编译时裁剪掉protobuf，因为整个tool工程没有用到protobuf，编译会去掉引用，然后动态编译就会出错
    class Table
    {
        public bool C;
        public bool S;
        public int Index;
        public Dictionary<string, HeadInfo> HeadInfos = new Dictionary<string, HeadInfo>();
    }
    public static partial class ExcelExporterCustom
    {
        private static string template;

        private const string ClientClassDir = "../Unity/Assets/Scripts/Model/Generate/Client/Config";
        // 服务端因为机器人的存在必须包含客户端所有配置，所以单独的c字段没有意义,单独的c就表示cs
        private const string ServerClassDir = "../Unity/Assets/Scripts/Model/Generate/Server/Config";

        private const string CSClassDir = "../Unity/Assets/Scripts/Model/Generate/ClientServer/Config";

        private const string excelDir = "../Unity/Assets/Config/ExcelCustom/";

        private const string jsonDir = "../Config/Json/{0}/{1}";

        private const string clientProtoDir = "../Unity/Assets/Bundles/Config";
        private const string serverProtoDir = "../Config/Excel/{0}/{1}";
        private static Assembly[] configAssemblies = new Assembly[3];

        private static Dictionary<string, Table> tables = new Dictionary<string, Table>();
        private static Dictionary<string, ExcelPackage> packages = new Dictionary<string, ExcelPackage>();

        private static Table GetTable(string protoName)
        {
            if (!tables.TryGetValue(protoName, out var table))
            {
                table = new Table();
                tables[protoName] = table;
            }

            return table;
        }

        public static ExcelPackage GetPackage(string filePath)
        {
            if (!packages.TryGetValue(filePath, out var package))
            {
                using Stream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                package = new ExcelPackage(stream);
                packages[filePath] = package;
            }

            return package;
        }

        public static void Export()
        {
            try
            {
                template = File.ReadAllText("Template.txt");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                //读取已导出过的excel文件md5
                Dictionary<string, string> fileMd5s = new Dictionary<string, string>();
                string md5File = $"{excelDir}/modified_md5.txt";
                if (File.Exists(md5File))
                {
                    string[] lines = File.ReadAllLines(md5File, Encoding.UTF8);
                    foreach (string line in lines)
                    {
                        string[] kv = line.Split('|');
                        fileMd5s[kv[0]] = kv[1];
                    }
                }

                List<string> existFiles = new List<string>();

                //读取所有excel文件
                List<string> files = FileHelper.GetAllFiles(excelDir);
                List<string> modifiedFiles = new List<string>();
                foreach (string path in files)
                {
                    string fileName = Path.GetFileName(path);
                    if (!fileName.EndsWith(".xlsx") || fileName.StartsWith("~$") || fileName.Contains("#"))
                    {
                        continue;
                    }
                    existFiles.Add(path);

                    var md5 = MD5.Create();
                    var bs = File.ReadAllBytes(path);
                    byte[] md5Bs = md5.ComputeHash(bs);
                    string curMd5Str = BitConverter.ToString(md5Bs);
                    fileMd5s.TryGetValue(path, out string oldMd5Str);
                    if (curMd5Str == oldMd5Str) continue;//excel文件无改动
                    //有改动，保存准备导出。
                    fileMd5s[path] = curMd5Str;
                    modifiedFiles.Add(path);
                    Log.Console($"准备到处文件：{path}");
                }
                //更新本次更改到md5文件。
                string newMd5Text = "";
                foreach (var kv in fileMd5s)
                {
                    if (existFiles.Contains(kv.Key))
                    {
                        newMd5Text += $"{kv.Key}|{kv.Value}{Environment.NewLine}";
                    }
                }
                File.WriteAllText(md5File, newMd5Text, Encoding.UTF8);

                foreach (string path in modifiedFiles)
                {
                    string fileName = Path.GetFileName(path);
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
                    string fileNameWithoutCS = fileNameWithoutExtension;
                    string cs = "cs";
                    if (fileNameWithoutExtension.Contains("@"))
                    {
                        string[] ss = fileNameWithoutExtension.Split("@");
                        fileNameWithoutCS = ss[0];
                        cs = ss[1];
                    }

                    if (cs == "")
                    {
                        cs = "cs";
                    }

                    ExcelPackage p = GetPackage(Path.GetFullPath(path));

                    string protoName = fileNameWithoutCS;
                    if (fileNameWithoutCS.Contains('_'))
                    {
                        protoName = fileNameWithoutCS.Substring(0, fileNameWithoutCS.LastIndexOf('_'));
                    }

                    Table table = GetTable(protoName);

                    if (cs.Contains("c"))
                    {
                        table.C = true;
                    }

                    if (cs.Contains("s"))
                    {
                        table.S = true;
                    }

                    ExportExcelClass(p, protoName, table);
                }

                foreach (var kv in tables)
                {
                    if (kv.Value.C)
                    {
                        ExportClass(kv.Key, kv.Value.HeadInfos, ConfigType.c);
                    }
                    if (kv.Value.S)
                    {
                        ExportClass(kv.Key, kv.Value.HeadInfos, ConfigType.s);
                    }
                    ExportClass(kv.Key, kv.Value.HeadInfos, ConfigType.cs);
                }

                // 动态编译生成的配置代码
                configAssemblies[(int)ConfigType.c] = DynamicBuild(ConfigType.c);
                configAssemblies[(int)ConfigType.s] = DynamicBuild(ConfigType.s);
                configAssemblies[(int)ConfigType.cs] = DynamicBuild(ConfigType.cs);

                //List<string> excels = FileHelper.GetAllFiles(excelDir, "*.xlsx");

                foreach (string path in modifiedFiles)
                {
                    ExportExcelToJson(path);
                }

                if (Directory.Exists(clientProtoDir))
                {
                    Directory.Delete(clientProtoDir, true);
                }
                FileHelper.CopyDirectory("../Config/Excel/c", clientProtoDir);

                Log.Console("Export Excel Sucess!");
            }
            catch (Exception e)
            {
                Log.Console(e.ToString());
            }
            finally
            {
                tables.Clear();
                foreach (var kv in packages)
                {
                    kv.Value.Dispose();
                }

                packages.Clear();
            }
        }

        private static void ExportExcelToJson(string path)
        {
            string dir = Path.GetDirectoryName(path);
            string relativePath = Path.GetRelativePath(excelDir, dir);
            string fileName = Path.GetFileName(path);
            if (!fileName.EndsWith(".xlsx") || fileName.StartsWith("~$") || fileName.Contains("#"))
            {
                return;
            }

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            string fileNameWithoutCS = fileNameWithoutExtension;
            string cs = "cs";
            if (fileNameWithoutExtension.Contains("@"))
            {
                string[] ss = fileNameWithoutExtension.Split("@");
                fileNameWithoutCS = ss[0];
                cs = ss[1];
            }

            if (cs == "")
            {
                cs = "cs";
            }

            string protoName = fileNameWithoutCS;
            if (fileNameWithoutCS.Contains('_'))
            {
                protoName = fileNameWithoutCS.Substring(0, fileNameWithoutCS.LastIndexOf('_'));
            }

            Table table = GetTable(protoName);

            ExcelPackage p = GetPackage(Path.GetFullPath(path));

            if (cs.Contains("c"))
            {
                ExportExcelJson(p, fileNameWithoutCS, table, ConfigType.c, relativePath);
                ExportExcelProtobuf(ConfigType.c, protoName, relativePath);
            }

            if (cs.Contains("s"))
            {
                ExportExcelJson(p, fileNameWithoutCS, table, ConfigType.s, relativePath);
                ExportExcelProtobuf(ConfigType.s, protoName, relativePath);
            }
            ExportExcelJson(p, fileNameWithoutCS, table, ConfigType.cs, relativePath);
            ExportExcelProtobuf(ConfigType.cs, protoName, relativePath);
        }

        private static string GetProtoDir(ConfigType configType, string relativeDir)
        {
            return string.Format(serverProtoDir, configType.ToString(), relativeDir);
        }

        private static Assembly GetAssembly(ConfigType configType)
        {
            return configAssemblies[(int)configType];
        }

        private static string GetClassDir(ConfigType configType)
        {
            return configType switch
            {
                ConfigType.c => ClientClassDir,
                ConfigType.s => ServerClassDir,
                _ => CSClassDir
            };
        }

        // 动态编译生成的cs代码
        private static Assembly DynamicBuild(ConfigType configType)
        {
            string classPath = GetClassDir(configType);
            List<SyntaxTree> syntaxTrees = new List<SyntaxTree>();
            List<string> protoNames = new List<string>();
            foreach (string classFile in Directory.GetFiles(classPath, "*.cs"))
            {
                protoNames.Add(Path.GetFileNameWithoutExtension(classFile));
                syntaxTrees.Add(CSharpSyntaxTree.ParseText(File.ReadAllText(classFile)));
            }

            List<PortableExecutableReference> references = new List<PortableExecutableReference>();
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
            foreach (Assembly assembly in assemblies)
            {
                try
                {
                    if (assembly.IsDynamic)
                    {
                        continue;
                    }

                    if (assembly.Location == "")
                    {
                        continue;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }

                PortableExecutableReference reference = MetadataReference.CreateFromFile(assembly.Location);
                references.Add(reference);
            }

            CSharpCompilation compilation = CSharpCompilation.Create(null,
                syntaxTrees.ToArray(),
                references.ToArray(),
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

            using MemoryStream memSteam = new MemoryStream();

            EmitResult emitResult = compilation.Emit(memSteam);
            if (!emitResult.Success)
            {
                StringBuilder stringBuilder = new StringBuilder();
                foreach (Diagnostic t in emitResult.Diagnostics)
                {
                    stringBuilder.Append($"{t.GetMessage()}\n");
                }

                throw new Exception($"dynamical build failed, 动态编译失败:\n{stringBuilder}");
            }

            memSteam.Seek(0, SeekOrigin.Begin);

            Assembly ass = Assembly.Load(memSteam.ToArray());
            return ass;
        }


        // 根据生成的类，把json转成protobuf
        private static void ExportExcelProtobuf(ConfigType configType, string protoName, string relativeDir)
        {
            string dir = GetProtoDir(configType, relativeDir);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            Assembly ass = GetAssembly(configType);
            Type type = ass.GetType($"ET.{protoName}Category");
            Type subType = ass.GetType($"ET.{protoName}");

            IMerge final = Activator.CreateInstance(type) as IMerge;

            string p = Path.Combine(string.Format(jsonDir, configType, relativeDir));
            string[] ss = Directory.GetFiles(p, $"{protoName}*.txt");
            List<string> jsonPaths = ss.ToList();

            jsonPaths.Sort();
            jsonPaths.Reverse();
            foreach (string jsonPath in jsonPaths)
            {
                string json = File.ReadAllText(jsonPath);
                json = $"{{\"dict\":{json}}}";
                try
                {
                    object deserialize = BsonSerializer.Deserialize(json, type);
                    final.Merge(deserialize);
                }
                catch (Exception e)
                {
                    throw new Exception($"json : {jsonPath} error", e);
                }
            }

            string path = Path.Combine(dir, $"{protoName}Category.bytes");

            using FileStream file = File.Create(path);
            file.Write(final.ToBson());
            file.Close();

            Log.Console($"Create bytes file : {path}");
        }
    }
}
