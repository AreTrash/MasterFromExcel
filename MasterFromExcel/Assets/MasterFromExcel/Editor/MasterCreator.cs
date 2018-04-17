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
            });

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
                try
                {
                    var path = MfeConst.ScriptableObjectScriptOutputDirectory + sc.MasterName + "Data.cs";
                    var columns = sc.Sheet.GetRow(0);
                    var types = sc.Sheet.GetRow(1);
                    var code = scriptableObjectCodeGenerator.GenerateCode(sc.MasterName, columns, types);
                    ImportScript(code, path);
                }
                catch (Exception)
                {
                    Debug.LogAssertion($"error occurred in {sc.MasterName}");
                    throw;
                }
            }
        }

        static void GenerateMasterInstallerScript(string miTemp, IEnumerable<SheetContext> sheetContexts)
        {
            var masterInstallerCodeGenerator = new MasterInstallerCodeGenerator(miTemp, MfeConst.MasterNamespace);

            try
            {
                var code = masterInstallerCodeGenerator.GenerateCode(sheetContexts.Select(sc => sc.MasterName));
                ImportScript(code, MfeConst.MasterInstallerOutputPath);
            }
            catch (Exception)
            {
                Debug.LogAssertion("error occurred in generating installer script");
                throw;
            }
        }

        static void GenerateMasterKeyValueObjectsScript(string mkvoTemp, IEnumerable<SheetContext> sheetContexts)
        {
            var masterKeyValueObjectsGenerator = new MasterKeyValueObjectsGenerator(mkvoTemp, MfeConst.MasterNamespace);

            try
            {
                var types = sheetContexts.SelectMany(sc => sc.Sheet.GetRow(1).Cells.Select(c => c.ToString()));
                var code = masterKeyValueObjectsGenerator.GenerateCode(types);
                ImportScript(code, MfeConst.MasterKeyValueObjectsOutputPath);
            }
            catch (Exception)
            {
                Debug.LogAssertion("error occurred in generating mkvo scripts");
                throw;
            }
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

            var key = ConvertibleTypeUtility.GetString(types.GetCell(0).StringCellValue);

            return scriptableObjectCodeTemplate
                .Replace("$Namespace$", @namespace)
                .Replace("$Master$", masterName)
                .Replace("$Key$", key)
                .Replace("$Columns$", fields.ToString().TrimEnd(Environment.NewLine))
                .Replace("$ResourcePath$", GetPathUnderResources());
        }

        string MakeField(string column, string type)
        {
            return type == "enum" ? 
                $"public {column} {column};" : 
                $"public {ConvertibleTypeUtility.GetString(type)} {column};";
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
    public struct $Key$
    {
        public string Value;
    }"
                .Replace("$Key$", ConvertibleTypeUtility.GetString(key));
        }
    }
}