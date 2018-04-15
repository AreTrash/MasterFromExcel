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
        //TODO Dry
        const string ScriptableObjectTemplate = @"Assets/Plugins/MasterFromExcel/Editor/MasterScriptableObjectTemplate.txt";
        const string MasterNamespace = @"Master";
        const string MasterExcelPath = @"../MasterData/";
        const string ScriptableObjectOutputPath = @"Assets/MasterFromExcel/";

        [MenuItem("Tools/MasterFromExcel/GenerateMasterScript")]
        static void GenerateScriptableObjectScript()
        {
            var excelPaths = GetExcelPaths();
            var soTemp = AssetDatabase.LoadAssetAtPath<TextAsset>(ScriptableObjectTemplate).text;
            var scriptableObjectCodeGenerator = new ScriptableObjectCodeGenerator(soTemp, MasterNamespace);

            foreach (var excelPath in excelPaths.Take(1))
            {
                var masterName = Path.GetFileNameWithoutExtension(excelPath).ToTopUpper();
                var sheet = WorkbookFactory.Create(excelPath).GetSheetAt(0);

                var path = ScriptableObjectOutputPath + masterName + "Data.cs";
                var code = scriptableObjectCodeGenerator.GenerateCode(masterName, sheet.GetRow(0), sheet.GetRow(1));
                WriteScriptableObjectScript(path, code);

                AssetDatabase.ImportAsset(path);
            }
        }

        static IEnumerable<string> GetExcelPaths()
        {
            return Directory.GetFiles(MasterExcelPath).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx"));
        }

        static void WriteScriptableObjectScript(string path, string code)
        {
            if (File.Exists(path) && code == File.ReadAllText(path))
            {
                //return;//変更がない場合はコンパイルを走らせないようにするため書き込まないのが理想
            }

            var directory = Path.GetDirectoryName(ScriptableObjectOutputPath);
            if (directory != null && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
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

        public ScriptableObjectCodeGenerator(string scriptableObjectCodeTemplate, string @namespace)
        {
            this.scriptableObjectCodeTemplate = scriptableObjectCodeTemplate;
            this.@namespace = @namespace;
        }

        public string GenerateCode(string masterName, IRow columns, IRow types)
        {
            var fields = new StringBuilder();

            for (var i = 0; i < columns.Count(); i++)
            {
                var column = columns.GetCell(i).StringCellValue.ToTopUpper();
                var type = types.GetCell(i).StringCellValue;
                fields.Append($"        {MakeField(column, type)}\n");
            }

            return scriptableObjectCodeTemplate
                .Replace("$Namespace$", @namespace)
                .Replace("$Master$", masterName)
                .Replace("$Columns$", fields.ToString().TrimEnd("\n"));
        }

        string MakeField(string column, string type)
        {
            if (ConvertibleTypes.Contains(type.Replace("[]", "")))
            {
                return $"public {type} {column};";
            }

            return $"public string {column};";
        }
    }
}
