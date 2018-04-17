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
        const string ScriptableObjectTemplateName = @"MasterScriptableObjectTemplate.txt";
        const string MasterInstallerTemplateName = @"MasterInstallerTemplate.txt";

        [MenuItem("Tools/MasterFromExcel/GenerateScript")]
        static void Generate()
        {
            var allAssetPaths = AssetDatabase.GetAllAssetPaths();
            var sotPath = allAssetPaths.First(s => s.EndsWith(ScriptableObjectTemplateName));
            var mitPath = allAssetPaths.First(s => s.EndsWith(MasterInstallerTemplateName));
            var excelPaths = Directory.GetFiles(MfeConst.MasterExcelDirectory).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx"));

            GenerateScriptableObjectScript(sotPath, excelPaths);
            GenerateMasterInstallerScript(mitPath, excelPaths);
        }

        static void GenerateScriptableObjectScript(string sotPath, IEnumerable<string> excelPaths)
        {
            var soTemp = AssetDatabase.LoadAssetAtPath<TextAsset>(sotPath).text;
            var scriptableObjectCodeGenerator = new ScriptableObjectCodeGenerator(
                soTemp, MfeConst.MasterNamespace, MfeConst.ScriptableObjectOutputDirectory
            );

            foreach (var excelPath in excelPaths)
            {
                var masterName = GetMasterName(excelPath);
                var sheet = WorkbookFactory.Create(excelPath).GetSheetAt(0);

                var path = MfeConst.ScriptableObjectScriptOutputDirectory + masterName + "Data.cs";
                var code = scriptableObjectCodeGenerator.GenerateCode(masterName, sheet.GetRow(0), sheet.GetRow(1));
                sheet.Workbook.Close();

                ImportScript(path, code);
            }
        }

        static void GenerateMasterInstallerScript(string mitPath, IEnumerable<string> excelPaths)
        {
            var miTemp = AssetDatabase.LoadAssetAtPath<TextAsset>(mitPath).text;
            var masterInstallerCodeGenerator = new MasterInstallerCodeGenerator(miTemp, MfeConst.MasterNamespace);
            var masterNames = excelPaths.Select(GetMasterName);
            var code = masterInstallerCodeGenerator.GenerateCode(masterNames);
            ImportScript(MfeConst.MasterInstallerOutputPath, code);
        }

        static string GetMasterName(string excelPath)
        {
            return Path.GetFileNameWithoutExtension(excelPath).ToTopUpper();
        }

        static void ImportScript(string path, string code)
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
        static readonly string[] ConvertibleTypes =
        {
            "int", "long", "float", "double", "bool", "string", "datetime",
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
                var type = types.GetCell(i).StringCellValue.ToLower();
                fields.Append($"        {MakeField(column, type)}{Environment.NewLine}");
            }

            return scriptableObjectCodeTemplate
                .Replace("$Namespace$", @namespace)
                .Replace("$Master$", masterName)
                .Replace("$Columns$", fields.ToString().TrimEnd(Environment.NewLine))
                .Replace("$ResourcePath$", GetPathUnderResources());
        }

        string MakeField(string column, string type)
        {
            if (type == "enum")
            {
                return $"public {column} {column};";
            }

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
            var bindings = string.Join(Environment.NewLine, masterNames.Select(mn => $"            {MakeBindings(mn)}"));

            return masterInstallerCodeTemplate
                .Replace("$Namespace$", @namespace)
                .Replace("$Bindings$", bindings);
        }

        string MakeBindings(string masterName)
        {
            return $"Container.Bind<I{masterName}Dao>().To<{masterName}Dao>().AsSingle();";
        }
    }
}