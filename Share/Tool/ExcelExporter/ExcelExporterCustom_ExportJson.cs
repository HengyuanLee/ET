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
        #region 导出json


        static void ExportExcelJson(ExcelPackage p, string name, Table table, ConfigType configType, string relativeDir)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append($"{{{Environment.NewLine}");
            foreach (ExcelWorksheet worksheet in p.Workbook.Worksheets)
            {
                if (worksheet.Name.StartsWith("#"))
                {
                    continue;
                }

                ExportSheetJson(worksheet, name, table.HeadInfos, configType, sb);
            }

            sb.Append($"}}");

            string dir = string.Format(jsonDir, configType.ToString(), relativeDir);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            string jsonPath = Path.Combine(dir, $"{name}.txt");
            using FileStream txt = new FileStream(jsonPath, FileMode.Create);
            using StreamWriter sw = new StreamWriter(txt);
            sw.Write(sb.ToString());
            sw.Dispose();
            sw.Close();
            Log.Console($"Create json file : {jsonPath}");
        }

        static void ExportSheetJson(ExcelWorksheet worksheet, string name,
                Dictionary<string, HeadInfo> classField, ConfigType configType, StringBuilder sb)
        {
            if (worksheet == null || worksheet.Dimension == null || worksheet.Dimension.End == null) return;

            string endStr = $",{Environment.NewLine}";

            string configTypeStr = configType.ToString();
            for (int row = 5; row <= worksheet.Dimension.End.Row; ++row)
            {
                string idValue = worksheet.Cells[row, 1].Text.Trim();
                if (idValue == "")
                {
                    continue;
                }
                sb.Append($"{Tab(1)}\"{idValue}\": {{{Environment.NewLine}{Tab(2)}\"_t\":\"{name}\",{Environment.NewLine}");
                for (int col = 1; col <= worksheet.Dimension.End.Column; ++col)
                {
                    string fieldName = worksheet.Cells[1, col].Text.Trim();
                    fieldName = fieldName.First().ToString().ToUpper() + fieldName.Substring(1);
                    if (!classField.ContainsKey(fieldName))
                    {
                        continue;
                    }

                    HeadInfo headInfo = classField[fieldName];

                    if (headInfo == null)
                    {
                        continue;
                    }

                    if (configType != ConfigType.cs && !headInfo.FieldCS.Contains(configTypeStr))
                    {
                        continue;
                    }

                    string fieldN = headInfo.FieldName;
                    if (fieldN.ToLower() == "id")
                    {
                        fieldN = "_id";
                    }

                    sb.Append($"{Tab(2)}\"{fieldN}\":{Convert(headInfo.FieldType, worksheet.Cells[row, col].Text.Trim())},{Environment.NewLine}");
                }
                sb.Replace(endStr, Environment.NewLine, sb.Length - endStr.Length, endStr.Length);
                sb.Append($"{Tab(1)}}},{Environment.NewLine}");
            }
            sb.Replace(endStr, Environment.NewLine, sb.Length-endStr.Length, endStr.Length);
        }
        private static string Tab(int num = 1)
        {
            string result = "";
            for (int i = 0; i < num; i++)
            {
                result += "\t";
            }
            return result;
        }

        private static string Convert(string type, string value)
        {
            switch (type)
            {
                case "uint[]":
                case "int[]":
                case "int32[]":
                case "long[]":
                    return $"[{value}]";
                case "string[]":
                case "int[][]":
                    return $"[{value}]";
                case "int":
                case "uint":
                case "int32":
                case "int64":
                case "long":
                case "float":
                case "double":
                    if (value == "")
                    {
                        return "0";
                    }

                    return value;
                case "string":
                    value = value.Replace("\\", "\\\\");
                    value = value.Replace("\"", "\\\"");
                    return $"\"{value}\"";
                default:
                    if (type.StartsWith("Dictionary"))
                    {
                        return GetDictValue(type, value);
                    }
                    throw new Exception($"不支持此类型: {type}");
            }
        }
        private static string GetDictValue(string type, string value)
        {
            string valueType = type.Replace("Dictionary<", string.Empty).Replace(">", string.Empty).Split(",")[1].Trim();

            string result = "";
            string[] kvStrs = value.Split(",");
            foreach (string kvStr in kvStrs)
            {
                string[] kvs = default;
                if (kvStr.Contains("="))
                {
                    kvs = kvStr.Split("=");
                }
                else if (kvStr.Contains(":"))
                {
                    kvs = kvStr.Split(":");
                }
                if (kvs.Length >= 2)
                {
                    string _key = kvs[0].Trim();
                    string _value = kvs[1].Trim();
                    result += $"{Tab(3)}\"{_key}\":{Convert(valueType, _value)},{Environment.NewLine}";
                }
            }
            result = $"{Environment.NewLine}{result.Remove(result.Length - 1 - Environment.NewLine.Length)}";
            result = $"{Environment.NewLine}{Tab(2)}{{{result}";
            result = $"{result}{Environment.NewLine}{Tab(2)}}}";
            return result;
        }

        #endregion
    }
}
