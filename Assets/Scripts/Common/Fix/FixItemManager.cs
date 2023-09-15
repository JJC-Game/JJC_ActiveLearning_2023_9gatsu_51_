using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

public class FixItemManager : MonoBehaviour
{
    // Start is called before the first frame update
    void Awake()
    {
        Load();
        DB_Disp();
    }

    // Update is called once per frame
    void Update()
    {
        
    }

    public void Load()
    {
        string path = "FixData/ItemFixData";

        TextAsset csvFile;
        fixItemDataList = new List<FixItemData>();

        {
            csvFile = Resources.Load<TextAsset>(path);
            StringReader reader = new StringReader(csvFile.text);

            while (reader.Peek() != -1)
            {
                string line = reader.ReadLine();
                string[] lineArray = line.Split(',');

                FixItemData currentRow = new FixItemData();
                currentRow.id = int.Parse(lineArray[0]);
                currentRow.name = lineArray[1];
                currentRow.imagePath = lineArray[2];

                fixItemDataList.Add(currentRow);
            }
        }
    }

    public FixItemData GetFixItemData(int itemId)
    {
        if (itemId >= GetFixItemNum())
        {
            return null;
        }
        return fixItemDataList[itemId];
    }

    public int GetFixItemNum()
    {
        return fixItemDataList.Count;
    }

    public int GetItemIdFromItemName(string itemName)
    {
        for (int i = 0; i < GetFixItemNum(); i++)
        {
            FixItemData data = GetFixItemData(i);
            if (data.name.Equals(itemName))
            {
                return i;
            }
        }
        return INVALID_ID;
    }

    public bool IsValidItemId(int itemId)
    {
        return 0 <= itemId && itemId < GetFixItemNum();
    }

    public void DB_Disp()
    {
        for (int i = 0; i < GetFixItemNum(); i++)
        {
            FixItemData data = GetFixItemData(i);
            Debug.Log(data.id + "," + data.name + "," + data.imagePath);
        }
    }

    public class FixItemData
    {
        public int id;
        public string name;
        public string imagePath;
    }

    List<FixItemData> fixItemDataList;
    public const int INVALID_ID = -1;
}
