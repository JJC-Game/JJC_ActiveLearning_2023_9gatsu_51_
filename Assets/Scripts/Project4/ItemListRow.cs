using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class ItemListRow : MonoBehaviour
{
    void Awake(){
        nameText = this.transform.Find("HLayout/ItemName").GetComponent<Text>();
        itemImage = this.transform.Find("HLayout/ItemImage").GetComponent<Image>();
        itemNum = this.transform.Find("HLayout/ItemNum").GetComponent<Text>();

        Font font = Resources.Load<Font>("Fonts/keifont");
        nameText.font = font;
        itemNum.font = font;
    }

    public void RefreshItemImage(int itemId){
        if (UserApplication.fixItemManager == null)
        {
            return;
        }
        if (UserApplication.fixItemManager.GetFixItemData(itemId) == null)
        {
            return;
        }
        nameText.text = UserApplication.fixItemManager.GetFixItemData(itemId).name;
        nameText.color = new Color(0, 0, 0);
        itemImage.sprite = Resources.Load<Sprite>(UserApplication.fixItemManager.GetFixItemData(itemId).imagePath);
        itemNum.text = UserApplication.p4_userDataManager.GetHasItemNum(itemId).ToString();
    }

    Text nameText;
    Image itemImage;
    Text itemNum;
}
