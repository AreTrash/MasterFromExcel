using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

namespace MasterFromExcel
{
    public static class MasterCreator
    {
        const string ScriptableObjectTemplateName = @"MasterScriptableObjectTemplate.txt";
        const string MasterInstallerTemplateName = @"MasterInstallerTemplate.txt";
        const string MasterKeyValueObjectsTemplateName = @"MasterKeyValueObjectsTemplate.txt";

        struct SheetContext
        {
            public string MasterName { get; }
            public ISheet Sheet { get; }

            public SheetContext(string masterName, ISheet sheet)
            {
                MasterName = masterName;
                Sheet = sheet;
            }
        }

        [MenuItem("Tools/MasterFromExcel/GenerateScript")]
        static void Generate()
        {
            var allAssetPaths = AssetDatabase.GetAllAssetPaths();
            var excelPaths = Directory.GetFiles(MfeConst.MasterExcelDirectory).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx"));
            var sheetContexts = excelPaths.Select(ep =>
            {
                var masterName = Path.GetFileNameWithoutExtension(ep).ToTopUpper();
                var sheet = WorkbookFactory.Create(ep).GetSheetAt(0);
                return new SheetContext(masterName, sheet);
            }).ToArray();

            var soTemp = GetTemplate(ScriptableObjectTemplateName, allAssetPaths);
            GenerateScriptableObjectScript(soTemp, sheetContexts);

            var miTemp = GetTemplate(MasterInstallerTemplateName, allAssetPaths);
            GenerateMasterInstallerScript(miTemp, sheetContexts);

            var mkvoTemp = GetTemplate(MasterKeyValueObjectsTemplateName, allAssetPaths);
            GenerateMasterKeyValueObjectsScript(mkvoTemp, sheetContexts);

            foreach (var sc in sheetContexts) sc.Sheet.Workbook.Close();
        }

        static string GetTemplate(string templateName, string[] allAssetPaths)
        {
            var path = allAssetPaths.Single(s => s.EndsWith(templateName));
            return AssetDatabase.LoadAssetAtPath<TextAsset>(path).text;
        }

        static void GenerateScriptableObjectScript(string soTemp, IEnumerable<SheetContext> sheetContexts)
        {
            var scriptableObjectCodeGenerator = new ScriptableObjectCodeGenerator(
                soTemp, MfeConst.MasterNamespace, MfeConst.ScriptableObjectOutputDirectory
            );

            foreach (var sc in sheetContexts)
            {
                var path = MfeConst.ScriptableObjectScriptOutputDirectory + sc.MasterName + "Data.cs";
                var columns = sc.Sheet.GetRow(0);
                var types = sc.Sheet.GetRow(1);
                var code = scriptableObjectCodeGenerator.GenerateCode(sc.MasterName, columns, types);
                ImportScript(code, path);
            }
        }

        static void GenerateMasterInstallerScript(string miTemp, IEnumerable<SheetContext> sheetContexts)
        {
            var masterInstallerCodeGenerator = new MasterInstallerCodeGenerator(miTemp, MfeConst.MasterNamespace);
            var code = masterInstallerCodeGenerator.GenerateCode(sheetContexts.Select(sc => sc.MasterName));
            ImportScript(code, MfeConst.MasterInstallerOutputPath);
        }

        static void GenerateMasterKeyValueObjectsScript(string mkvoTemp, IEnumerable<SheetContext> sheetContexts)
        {
            var masterKeyValueObjectsGenerator = new MasterKeyValueObjectsGenerator(mkvoTemp, MfeConst.MasterNamespace);
            var types = sheetContexts.SelectMany(sc => sc.Sheet.GetRow(1).Cells.Select(c => c.ToString()));
            var code = masterKeyValueObjectsGenerator.GenerateCode(types);
            ImportScript(code, MfeConst.MasterKeyValueObjectsOutputPath);
        }

        static void ImportScript(string code, string path)
        {
            if (File.Exists(path) && code == File.ReadAllText(path))
            {
                Debug.Log($"no change in {path}");
                return;
            }

            var directory = Path.GetDirectoryName(path);

            if (directory != null && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(path, code);
            AssetDatabase.ImportAsset(path);
            Debug.Log($"import {path}");
        }
    }

    public static class ConvertibleTypeUtility
    {
        static readonly IEnumerable<string> ConvertibleTypes = new[]
        {
            "int", "long", "float", "double", "bool", "string", "datetime", "enum",
        };

        public static bool IsDefined(string type)
        {
            return ConvertibleTypes.Contains(Regex.Replace(type.ToLower(), @"\[\]$", ""));
        }
    }

    public class ScriptableObjectCodeGenerator
    {
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
                var type = types.GetCell(i).StringCellValue.ToLower();
                fields.Append($"        {MakeField(column, type)}{Environment.NewLine}");
            }

            var keyType = types.GetCell(0).StringCellValue;
            var key = ConvertibleTypeUtility.IsDefined(keyType) ? keyType.ToLower() : $"{keyType.ToTopUpper()}Key";

            return scriptableObjectCodeTemplate
                .Replace("$Namespace$", @namespace)
                .Replace("$Master$", masterName)
                .Replace("$Key$", key)
                .Replace("$Columns$", fields.ToString().TrimEnd(Environment.NewLine))
                .Replace("$ResourcePath$", GetPathUnderResources());
        }

        string MakeField(string column, string type)
        {
            if (type == "enum")
            {
                return $"public {column} {column};";
            }

            if (ConvertibleTypeUtility.IsDefined(type))
            {
                return $"public {type} {column};";
            }

            return $"public {type.ToTopUpper()}Key {column};";
        }

        string GetPathUnderResources()
        {
            const string Resources = "Resources/";
            var index = scriptableObjectPath.IndexOf(Resources, StringComparison.Ordinal) + Resources.Length;
            return scriptableObjectPath.Substring(index);
        }
    }

    public class MasterInstallerCodeGenerator
    {
        readonly string masterInstallerCodeTemplate;
        readonly string @namespace;

        public MasterInstallerCodeGenerator(string masterInstallerCodeTemplate, string @namespace)
        {
            this.masterInstallerCodeTemplate = masterInstallerCodeTemplate;
            this.@namespace = @namespace;
        }

        public string GenerateCode(IEnumerable<string> masterNames)
        {
            var bindings = string.Join(Environment.NewLine, masterNames.Select(mn => $"            {MakeBinding(mn)}"));

            return masterInstallerCodeTemplate
                .Replace("$Namespace$", @namespace)
                .Replace("$Bindings$", bindings);
        }

        string MakeBinding(string masterName)
        {
            return $"Container.Bind<I{masterName}Dao>().To<{masterName}Dao>().AsSingle();";
        }
    }

    public class MasterKeyValueObjectsGenerator
    {
        readonly string masterKeyValueObjectsCodeTemplate;
        readonly string @namespace;

        public MasterKeyValueObjectsGenerator(string masterKeyValueObjectsCodeTemplate, string @namespace)
        {
            this.masterKeyValueObjectsCodeTemplate = masterKeyValueObjectsCodeTemplate;
            this.@namespace = @namespace;
        }

        public string GenerateCode(IEnumerable<string> types)
        {
            var keys = types.Distinct().Where(t => !ConvertibleTypeUtility.IsDefined(t));

            var kvos = string
                .Join(Environment.NewLine + Environment.NewLine, keys.Select(MakeKvo))
                .TrimEnd(Environment.NewLine);

            return masterKeyValueObjectsCodeTemplate
                .Replace("$Namespace$", @namespace)
                .Replace("$KeyValueObjects$", kvos);
        }

        string MakeKvo(string key)
        {
            return @"    [Serializable]
    public struct $Key$Key
    {
        public string Value;
    }"
                .Replace("$Key$", key.ToTopUpper());
        }
    }
}