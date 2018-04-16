using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

namespace MasterFromExcel
{
    public class MasterCreator
    {
        const string ScriptableObjectTemplate = @"Assets/Plugins/MasterFromExcel/Editor/MasterScriptableObjectTemplate.txt";

        [MenuItem("Tools/MasterFromExcel/GenerateScript")]
        static void GenerateScriptableObjectScript()
        {
            var excelPaths = GetExcelPaths();
            var soTemp = AssetDatabase.LoadAssetAtPath<TextAsset>(ScriptableObjectTemplate).text;
            var scriptableObjectCodeGenerator = new ScriptableObjectCodeGenerator(
                soTemp, MfeConst.MasterNamespace, MfeConst.ScriptableObjectOutputPath
            );

            foreach (var excelPath in excelPaths)
            {
                var masterName = Path.GetFileNameWithoutExtension(excelPath).ToTopUpper();
                var sheet = WorkbookFactory.Create(excelPath).GetSheetAt(0);

                var path = MfeConst.ScriptableObjectScriptOutputPath + masterName + "Data.cs";
                var code = scriptableObjectCodeGenerator.GenerateCode(masterName, sheet.GetRow(0), sheet.GetRow(1));
                sheet.Workbook.Close();

                WriteScriptableObjectScript(path, code);
                AssetDatabase.ImportAsset(path);
                Debug.Log($"import {path}");
            }
        }

        static IEnumerable<string> GetExcelPaths()
        {
            return Directory.GetFiles(MfeConst.MasterExcelPath).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx"));
        }

        static void WriteScriptableObjectScript(string path, string code)
        {
            if (File.Exists(path) && code == File.ReadAllText(path))
            {
                //return;//変更がない場合はコンパイルを走らせないようにするため書き込まないのが理想
            }

            if (!Directory.Exists(MfeConst.ScriptableObjectScriptOutputPath))
            {
                Directory.CreateDirectory(MfeConst.ScriptableObjectScriptOutputPath);
            }

            File.WriteAllText(path, code);
        }
    }

    public class ScriptableObjectCodeGenerator
    {
        static readonly string[] ConvertibleTypes =
        {
            "int", "long", "float", "double", "bool", "string", "DateTime",
        };

        readonly string scriptableObjectCodeTemplate;
        readonly string @namespace;
        readonly string scriptableObjectPath;

        public ScriptableObjectCodeGenerator(string scriptableObjectCodeTemplate, string @namespace, string scriptableObjectPath)
        {
            this.scriptableObjectCodeTemplate = scriptableObjectCodeTemplate;
            this.@namespace = @namespace;
            this.scriptableObjectPath = scriptableObjectPath;
        }

        public string GenerateCode(string masterName, IRow columns, IRow types)
        {
            var fields = new StringBuilder();

            for (var i = 0; i < columns.Count(); i++)
            {
                var column = columns.GetCell(i).StringCellValue.ToTopUpper();
                var type = types.GetCell(i).StringCellValue;
                fields.Append($"        {MakeField(column, type)}{Environment.NewLine}");
            }

            return scriptableObjectCodeTemplate
                .Replace("$ResourcePath$", GetPathUnderResources())
                .Replace("$Namespace$", @namespace)
                .Replace("$Master$", masterName)
                .Replace("$Columns$", fields.ToString().TrimEnd(Environment.NewLine));
        }

        string MakeField(string column, string type)
        {
            if (ConvertibleTypes.Contains(type.Replace("[]", "")))
            {
                return $"public {type} {column};";
            }

            return $"public string {column};";
        }

        string GetPathUnderResources()
        {
            const string Resources = "Resources/";
            var index = scriptableObjectPath.IndexOf(Resources, StringComparison.Ordinal) + Resources.Length;
            return scriptableObjectPath.Substring(index);
        }
    }
}