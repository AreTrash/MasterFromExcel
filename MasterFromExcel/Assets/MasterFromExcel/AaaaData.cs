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
    }

    public interface IAaaaGetter
    {
        Aaaa[] GetAll();
        Aaaa Get(string key);
    }

    public class AaaaGetter : IAaaaGetter
    {
        readonly Dictionary<string, Aaaa> dataDic = new Dictionary<string, Aaaa>();

        public AaaaGetter()
        {
            throw new NotImplementedException();
        }

        public Aaaa[] GetAll()
        {
            throw new NotImplementedException();
        }

        public Aaaa Get(string key)
        {
            throw new NotImplementedException();
        }
    }
}