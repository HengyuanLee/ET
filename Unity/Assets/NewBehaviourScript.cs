using MongoDB.Bson.Serialization;
using MongoDB.Bson;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UnityEngine;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson.Serialization.Options;
using UnityEditor;

namespace ET
{
    public class CharacterSingleton : ET.Singleton<CharacterSingleton>
    {
        [BsonDictionaryOptions(DictionaryRepresentation.ArrayOfArrays)]
        public Dictionary<int, Character> dict;
    }

    public class NewBehaviourScript : MonoBehaviour
    {
        public TextAsset jsonAsset;
        [ContextMenu("Do it !")]
        void Start()
        {
            FileStream file = null;
            object deserialize = null;
            string bsPath = $"Assets/{nameof(CharacterCategory)}.bytes";
            string json = $"{{\"dict\":{jsonAsset.text}}}";
            deserialize = BsonSerializer.Deserialize<CharacterSingleton>(json);
            file = File.Create(bsPath);

            file.Write(deserialize.ToBson());
            file.Close();

            byte[] bs = AssetDatabase.LoadAssetAtPath<TextAsset>(bsPath).bytes;

            //LoadOneInThread(typeof(CharacterSingleton), bs);
            var cs = BsonSerializer.Deserialize<CharacterSingleton>(bs);

            foreach (var kv in cs.dict)
            {
                Debug.Log($"------------------->  <----- {kv.Value.Name}");
                foreach (var kk in kv.Value.Attr)
                {
                    Debug.Log($"{kk.Key} : {kk.Value}");
                }
            }
        }
        private void LoadOneInThread(Type configType, byte[] oneConfigBytes)
        {
            object category = MongoHelper.Deserialize(configType, oneConfigBytes, 0, oneConfigBytes.Length);

            lock (this)
            {
                ASingleton singleton = category as ASingleton;
                World.Instance.AddSingleton(singleton);
            }
        }

    }
    [Config]
    public partial class CharacterCategory : Singleton<CharacterCategory>, IMerge
    {
        [BsonElement]
        [BsonDictionaryOptions(DictionaryRepresentation.ArrayOfArrays)]
        private Dictionary<int, Character> dict = new();

        public void Merge(object o)
        {
            CharacterCategory s = o as CharacterCategory;
            foreach (var kv in s.dict)
            {
                this.dict.Add(kv.Key, kv.Value);
            }
        }

        public Character Get(int id)
        {
            this.dict.TryGetValue(id, out Character item);

            if (item == null)
            {
                throw new Exception($"�����Ҳ��������ñ���: {nameof(Character)}������id: {id}");
            }

            return item;
        }

        public bool Contain(int id)
        {
            return this.dict.ContainsKey(id);
        }

        public Dictionary<int, Character> GetAll()
        {
            return this.dict;
        }

        public Character GetOne()
        {
            if (this.dict == null || this.dict.Count <= 0)
            {
                return null;
            }
            return this.dict.Values.GetEnumerator().Current;
        }
    }
    public partial class Character : ProtoObject, IConfig
    {
        /// <summary>id</summary>
        public int Id { get; set; }
        /// <summary>����</summary>
        public string Name { get; set; }
        /// <summary>ְҵ����(ö��)</summary>
        public int Pro { get; set; }
        /// <summary>������(�ṹ��)</summary>
        public Dictionary<string, int> Spawn_point { get; set; }
        /// <summary>����(����)</summary>
        public int[] Weapon_id { get; set; }
        /// <summary>����(�ֵ�)</summary>
        public Dictionary<string, int> Attr { get; set; }
        /// <summary>����(�ֵ�)</summary>
        [BsonDictionaryOptions(DictionaryRepresentation.ArrayOfArrays)]
        public Dictionary<int, string> T1 { get; set; }

    }
}
