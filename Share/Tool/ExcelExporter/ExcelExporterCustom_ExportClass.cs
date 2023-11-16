using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ET.ExcelTool
{
    public static partial class ExcelExporterCustom
    {
        #region 导出class

        static void ExportExcelClass(ExcelPackage p, string name, Table table)
        {
            foreach (ExcelWorksheet worksheet in p.Workbook.Worksheets)
            {
                if (worksheet.Name.StartsWith("#") || worksheet.Name.Trim().ToLower() == "#alias")
                {
                    continue;
                }
                ExportSheetClass(worksheet, table);
            }
        }

        static void ExportSheetClass(ExcelWorksheet worksheet, Table table)
        {
            if (worksheet.Dimension == null || worksheet.Dimension.End == null) { return; }
            const int row = 0;
            for (int col = 1; col <= worksheet.Dimension.End.Column; ++col)
            {
                //行1：字段名称
                string fieldName = worksheet.Cells[row + 1, col].Text.Trim();
                //首字母大写
                fieldName = fieldName.First().ToString().ToUpper() + fieldName.Substring(1);
                if (fieldName == "" || fieldName.StartsWith("#"))
                {
                    continue;
                }

                if (table.HeadInfos.ContainsKey(fieldName))
                {
                    continue;
                }

                //行3：自定义配置
                Dictionary<string, string> configs = new Dictionary<string, string>();
                string configStr = worksheet.Cells[row + 3, col].Text.Trim();
                string[] configStrKVs = configStr.Split(',');
                foreach (string kvStr in configStrKVs)
                {
                    string[] kv = kvStr.Split(":");
                    if (kv.Length >= 2)
                    {
                        string key = kv[0].Trim();
                        string value = kv[1].Trim();
                        configs.Add(key, value);
                    }
                }
                //行2：字段类型
                string fieldType = worksheet.Cells[row + 2, col].Text.Trim();
                //行4：字段描述
                string fieldDesc = worksheet.Cells[row + 4, col].Text.Trim();

                table.HeadInfos[fieldName] = new HeadInfo(fieldDesc, fieldName, fieldType, ++table.Index, configs);
            }
        }

        static void ExportClass(string protoName, Dictionary<string, HeadInfo> classField, ConfigType configType)
        {
            string dir = GetClassDir(configType);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            string exportPath = Path.Combine(dir, $"{protoName}.cs");
            Log.Console($"Create c# file ：{exportPath}");
            using FileStream txt = new FileStream(exportPath, FileMode.Create);
            using StreamWriter sw = new StreamWriter(txt);

            StringBuilder sb = new StringBuilder();
            foreach ((string _, HeadInfo headInfo) in classField)
            {
                if (headInfo == null)
                {
                    continue;
                }

                if (configType != ConfigType.cs && !headInfo.FieldCS.Contains(configType.ToString()))
                {
                    continue;
                }

                string fieldType = headInfo.FieldType;
                bool needDictAttri = false;
                if (fieldType.StartsWith("Dictionary"))
                {
                    string[] kv = fieldType.Replace("Dictionary<", string.Empty).Replace(">", string.Empty).Split(",");
                    if (kv.Length > 0)
                    {
                        needDictAttri = !kv[0].Trim().Equals("string"); 
                    }
                }
                sb.Append($"\t\t/// <summary>{headInfo.FieldDesc}</summary>{Environment.NewLine}");
                //字典以非string为key时，Bson需要加标注。
                if (needDictAttri) sb.Append($"\t\t[BsonDictionaryOptions(DictionaryRepresentation.ArrayOfArrays)]{Environment.NewLine}");
                sb.Append($"\t\tpublic {fieldType} {headInfo.FieldName} {{ get; set; }}{Environment.NewLine}");
            }

            string content = template.Replace("(ConfigName)", protoName).Replace(("(Fields)"), sb.ToString());
            sw.Write(content);
            sw.Dispose();
            sw.Close();
        }

        #endregion
    }
}
