using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class P4_PHPConnectManager : PHPConnectManager
{
    public void CallShowHasItem(int userId)
    {
        calledUserId = userId;
        string phpFileName = "show_has_item/" + userId.ToString();
        CallPHPConnection(phpFileName, () => RefreshItemListUI());
    }

    private void RefreshItemListUI()
    {
        UserApplication.p4_userDataManager.SetHasItemNum(phpConnectResultText);
    }

    public void CallPlayItemGacha(int userId)
    {
        calledUserId = userId;
        string phpFileName = "play_item_gacha/" + userId.ToString();
        CallPHPConnection(phpFileName, () => RefreshItemGachaPerformUI());
    }
    private void RefreshItemGachaPerformUI()
    {
        GameObject obj = GameObject.Find("GachaManager");
        if (obj != null)
        {
            ItemGachaManager gachaManager = obj.GetComponent<ItemGachaManager>();
            gachaManager.DrawPerform(phpConnectResultText);
        }
        CallShowHasItem(calledUserId);
    }

}
