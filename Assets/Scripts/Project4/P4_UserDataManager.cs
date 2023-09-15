using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class P4_UserDataManager : P3_UserDataManager
{
    protected int[] hasItemNum;
    protected int itemIdNum;
    const int DEFAULT_ITEM_ID_NUM = 0;

    new protected void Awake()
    {
        base.Awake();

        itemIdNum = DEFAULT_ITEM_ID_NUM;
    }

    // Start is called before the first frame update
    void Start()
    {
        UserApplication.p4_phpConnectManager.CallShowHasItem(userId);
    }

    // Update is called once per frame
    void Update()
    {
        
    }

    public void SetHasItemNum(string inputText)
    {
        string[] textArray = inputText.Split(",");
        Debug.Log("入力"+inputText+"現在のアイテム数"+itemIdNum);
        if(itemIdNum == DEFAULT_ITEM_ID_NUM){
            Debug.Log("リスト配列を生成");
            CreateHasItemNumArray(textArray.Length - 1);
        }

        if (textArray.Length - 1 != itemIdNum)
        {
            Debug.LogAssertion("not match GetItem");
            return;
        }

        for(int i = 0; i < itemIdNum; i++){
            hasItemNum[i] = int.Parse(textArray[i]);
        }

        UserApplication.itemListRenderer.RefreshList();
    }

    public int GetHasItemNum(int itemId){
        return hasItemNum[itemId];
    }

    public int GetItemIdNum(){
        return itemIdNum;
    }

    void CreateHasItemNumArray(int inputItemIdNum){
        itemIdNum = inputItemIdNum;
        hasItemNum = new int[itemIdNum];
        for (int i = 0; i < itemIdNum; i++)
        {
            hasItemNum[i] = 0;
        }

        UserApplication.itemListRenderer.CreateHasItemNumArray(inputItemIdNum);
    }
}
