﻿using OfficeOpenXml;
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
                if (worksheet.Name.StartsWith("#"))
                {
                    continue;
                }
                ExportSheetClass(worksheet, table);
            }
        }

        static void ExportSheetClass(ExcelWorksheet worksheet, Table table)
        {
            const int row = 0;
            for (int col = 1; col <= worksheet.Dimension.End.Column; ++col)
            {
                //行1：字段名称
                string fieldName = worksheet.Cells[row + 1, col].Text.Trim();
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

                sb.Append($"\t\t/// <summary>{headInfo.FieldDesc}</summary>\n");
                string fieldType = headInfo.FieldType;
                sb.Append($"\t\tpublic {fieldType} {headInfo.FieldName} {{ get; set; }}\n");
            }

            string content = template.Replace("(ConfigName)", protoName).Replace(("(Fields)"), sb.ToString());
            sw.Write(content);
        }

        #endregion
    }
}
