using System;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

namespace MasterFromExcel
{
    public class MasterInjector : AssetPostprocessor
    {
        //TODO Dry
        const string ScriptableObjectOutputPath = @"Assets/MasterFromExcel/";
        const string MasterExcelPath = @"../MasterData/";
        const string MasterNamespace = @"Master";

        static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
        {
            var masterScriptPaths = importedAssets
                .Where(s => s.StartsWith(ScriptableObjectOutputPath))
                .Where(s => Path.GetExtension(s) == ".cs");

            foreach (var path in masterScriptPaths)
            {
                var asset = AssetDatabase.LoadAssetAtPath<MonoScript>(path);

                var scriptableObject = GenerateScriptableObject(asset.name);
                if (!scriptableObject) continue;

                AssetDatabase.CreateAsset(scriptableObject, ScriptableObjectOutputPath + asset.name + ".asset");
            }
        }

        static ScriptableObject GenerateScriptableObject(string name)
        {
            var nameWithoutData = name.EndsWith("Data") ? name.Substring(0, name.Length - "Data".Length) : name;

            ISheet sheet;
            var scriptableObject = TryGetSheet(nameWithoutData, out sheet) ? ScriptableObject.CreateInstance(name) : null;
            if (!scriptableObject) return null;

            var data = scriptableObject.GetType().GetField("Data").GetValue(scriptableObject);
            var add = data.GetType().GetMethod("Add", BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public);
            var valueGen = new ValueGenerator(sheet.GetRow(0), sheet.GetRow(1));

            foreach (IRow row in sheet)
            {
                if (row.RowNum <= 1) continue;
                add.Invoke(data, new[] { GetDto(row, valueGen, nameWithoutData) });
            }

            return scriptableObject;
        }

        static object GetDto(IRow row, ValueGenerator valueGen, string dtoName)
        {
            var dto = Activator.CreateInstance("Assembly-CSharp", MasterNamespace + "." + dtoName).Unwrap();

            foreach (var tup in row.Select((cell, i) => new { cell, i }))
            {
                var value = valueGen.GetValue(tup.i, tup.cell);
                var field = dto.GetType().GetField(valueGen.GetField(tup.i));
                field.SetValue(dto, value);
            }

            return dto;
        }

        static bool TryGetSheet(string masterName, out ISheet sheet)
        {
            foreach (var ee in new[] { ".xlsx", ".xls" })
            {
                var path = MasterExcelPath + masterName + ee;
                if (!File.Exists(path)) continue;
                sheet = WorkbookFactory.Create(path).GetSheetAt(0);
                return true;
            }

            sheet = null;
            return false;
        }
    }

    public class ValueGenerator
    {
        static readonly char[] Separator = {',', '/'};
        
        readonly IRow columns;
        readonly IRow types;

        public ValueGenerator(IRow columns, IRow types)
        {
            this.columns = columns;
            this.types = types;
        }

        public string GetField(int index)
        {
            return columns.Cells[index].StringCellValue.ToTopUpper();
        }

        public object GetValue(int index, ICell valueCell)
        {
            switch (types.Cells[index].StringCellValue)
            {
                case "int": return (int)valueCell.NumericCellValue;
                case "long": return (long)valueCell.NumericCellValue;
                case "float": return (float)valueCell.NumericCellValue;
                case "double": return valueCell.NumericCellValue;
                case "bool": return valueCell.BooleanCellValue;
                case "string": return valueCell.StringCellValue;
                case "DateTime": return valueCell.DateCellValue;

                case "int[]": return GetArrayValue(valueCell, int.Parse);
                case "long[]": return GetArrayValue(valueCell, long.Parse);
                case "float[]": return GetArrayValue(valueCell, float.Parse);
                case "double[]": return GetArrayValue(valueCell, double.Parse);
                case "bool[]": return GetArrayValue(valueCell, bool.Parse);
                case "string[]": return GetArrayValue(valueCell, s => s);
                case "DateTime[]": return GetArrayValue(valueCell, DateTime.Parse);

                default: return valueCell.StringCellValue;
            }
        }

        T[] GetArrayValue<T>(ICell valueCell, Func<string, T> parser)
        {
            return valueCell.StringCellValue
                .Split(Separator, StringSplitOptions.RemoveEmptyEntries)
                .Select(parser)
                .ToArray();
        }
    }
}
