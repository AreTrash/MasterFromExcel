using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using NPOI.SS.UserModel;
using UniRx;
using UnityEditor;
using UnityEditor.Callbacks;
using UnityEngine;

namespace MasterFromExcel
{
    public static class MasterInjector
    {
        [MenuItem("Tools/MasterFromExcel/InjectToScriptableObject")]
        static void CreateScriptableObject()
        {
            var scriptPaths = Directory
                .GetFiles(MfeConst.ScriptableObjectScriptOutputPath)
                .Where(p => Path.GetExtension(p) == ".cs");

            foreach (var scriptPath in scriptPaths)
            {
                var asset = AssetDatabase.LoadAssetAtPath<MonoScript>(scriptPath);
                try
                {
                    var scriptableObject = GenerateScriptableObject(asset.name);
                    if (!scriptableObject) return;

                    var path = MfeConst.ScriptableObjectOutputPath + asset.name + ".asset";
                    WriteScriptableObject(scriptableObject, path);
                    Debug.Log($"inject to {path}");
                }
                catch (Exception)
                {
                    Debug.LogAssertion($"error occurred in {asset.name}");
                    throw;
                }
            }
        }

        static void WriteScriptableObject(ScriptableObject scriptableObject, string path)
        {
            if (!Directory.Exists(MfeConst.ScriptableObjectOutputPath))
            {
                Directory.CreateDirectory(MfeConst.ScriptableObjectOutputPath);
            }

            //TODO タイムスタンプ毎回変わるけどいい？
            AssetDatabase.CreateAsset(scriptableObject, path);
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

            sheet.Workbook.Close();
            return scriptableObject;
        }

        static object GetDto(IRow row, ValueGenerator valueGen, string dtoName)
        {
            var dto = Activator.CreateInstance("Assembly-CSharp", MfeConst.MasterNamespace + "." + dtoName).Unwrap();

            for (var i = 0; i < valueGen.ColumnCount; i++)
            {
                var cell = row.GetCell(i);
                var value = valueGen.GetValue(i, cell);
                var field = dto.GetType().GetField(valueGen.GetField(i));
                field.SetValue(dto, value);
            }

            return dto;
        }

        static bool TryGetSheet(string masterName, out ISheet sheet)
        {
            foreach (var xlsExs in new[] { ".xlsx", ".xls" })
            {
                var path = MfeConst.MasterExcelPath + masterName + xlsExs;
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

        public int ColumnCount => columns.Cells.Count;

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
            if (valueCell == null) return null;

            switch (types.Cells[index].StringCellValue.ToLower())
            {
                case "int":
                    return (int)valueCell.NumericCellValue;
                case "long":
                    return (long)valueCell.NumericCellValue;
                case "float":
                    return (float)valueCell.NumericCellValue;
                case "double":
                    return valueCell.NumericCellValue;
                case "bool":
                    return valueCell.BooleanCellValue;
                case "string":
                    return valueCell.ToString();
                case "datetime":
                    return valueCell.DateCellValue;

                case "int[]":
                    return GetArrayValue(valueCell, int.Parse);
                case "long[]":
                    return GetArrayValue(valueCell, long.Parse);
                case "float[]":
                    return GetArrayValue(valueCell, float.Parse);
                case "double[]":
                    return GetArrayValue(valueCell, double.Parse);
                case "bool[]":
                    return GetArrayValue(valueCell, bool.Parse);
                case "string[]":
                    return GetArrayValue(valueCell, s => s);
                case "datetime[]":
                    return GetArrayValue(valueCell, DateTime.Parse);

                default:
                    return valueCell.ToString();
            }
        }

        T[] GetArrayValue<T>(ICell valueCell, Func<string, T> parser)
        {
            return valueCell.StringCellValue.Split(Separator)
                .Select(s => s == "" ? default(T) : parser(s))
                .ToArray();
        }
    }
}
