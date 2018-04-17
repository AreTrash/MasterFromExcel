/////////////////////////////////////////////////
//自動生成ファイルです！直接編集しないでください！//
/////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using UnityEngine;

namespace Master
{
    public class AaaaData : ScriptableObject
    {
        public List<Aaaa> Data = new List<Aaaa>();
    }

    [Serializable]
    public class Aaaa
    {
        public string Key;
        public int Arg;
        public string Effect;
        public float FloatTest;
        public double DoubleTest;
        public bool BoolTest;
        public string[] ArrayStringTest;
        public int[] ArrayIntTest;
        public TestEnum TestEnum;
    }

    public interface IAaaaDao
    {
        IEnumerable<Aaaa> GetAll();
        Aaaa Get(string key);
    }

    public class AaaaDao : IAaaaDao
    {
        readonly Dictionary<string, Aaaa> dataDic;

        public AaaaDao()
        {
            var AaaaObject = Resources.Load<AaaaData>("Master/Aaaa.asset");
            dataDic = AaaaObject.Data.ToDictionary(d => d.Key);
            Resources.UnloadAsset(AaaaObject);
        }

        public IEnumerable<Aaaa> GetAll()
        {
            return dataDic.Values;
        }

        public Aaaa Get(string key)
        {
            return dataDic[key];
        }
    }
}