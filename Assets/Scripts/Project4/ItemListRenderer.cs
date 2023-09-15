using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class ItemListRenderer : MonoBehaviour
{
    void Awake()
    {
        mouseHold = false;

        itemListRectTrans = this.transform.Find("RectMask/VLayout").transform as RectTransform;
        currentListPosition = itemListRectTrans.anchoredPosition;

        itemListRowNum = ITEM_ROW_INVALID;
    }
    // Start is called before the first frame update
    void Start()
    {
        RefreshList();
    }

    // Update is called once per frame
    void Update()
    {
        if (mouseHold)
        {
            Vector2 currentMousePosition = Input.mousePosition;
            Vector2 diffPosition = currentMousePosition - holdStartMousePosition;

            float yPosition = currentListPosition.y + diffPosition.y;

            if (yPosition < MIN_Y_POSITION)
            {
                yPosition = MIN_Y_POSITION;
            }
            if (yPosition > MAX_Y_POSITION)
            {
                yPosition = MAX_Y_POSITION;
            }

            itemListRectTrans.anchoredPosition = new Vector2(currentListPosition.x, yPosition);
        }
    }

    public void OnMouseButtonDown()
    {
        mouseHold = true;
        holdStartMousePosition = Input.mousePosition;
    }

    public void OnMouseButtonUp()
    {
        mouseHold = false;

        currentListPosition = itemListRectTrans.anchoredPosition;
    }

    virtual public void RefreshList()
    {
        for (int itemId = 0; itemId < itemListRowNum; itemId++)
        {
            Transform trans = this.transform.Find("RectMask/VLayout/ItemListRow" + itemId.ToString());
            ItemListRow itemListRow = trans.GetComponent<ItemListRow>();

            itemListRow.RefreshItemImage(itemId);
        }
    }

    public void CreateHasItemNumArray(int inputItemIdNum){
        Debug.Log("アイテム数"+inputItemIdNum.ToString()+"でリストを初期化");

        itemListRowNum = inputItemIdNum;
        MAX_Y_POSITION = (itemListRowNum - 3) * 90;

        GameObject itemListRowPrefab = Resources.Load<GameObject>("Prefabs/Item/ItemListRow");
        for (int i = 0; i < itemListRowNum; i++)
        {
            GameObject itemListRowInstance = Instantiate(itemListRowPrefab);
            itemListRowInstance.transform.name = "ItemListRow" + i.ToString();
            itemListRowInstance.transform.SetParent(itemListRectTrans);
        }
    }    

    bool mouseHold;
    Vector2 holdStartMousePosition;
    Vector2 currentListPosition;
    RectTransform itemListRectTrans;
    int itemListRowNum;
    const int ITEM_ROW_INVALID = 0;

    const int CELL_SIZE = 100;

    const int MIN_Y_POSITION = 0;
    int MAX_Y_POSITION = 0;

    const int MIN_CHARA_ID = 0;
    const int MAX_CHARA_ID = 31;
}
