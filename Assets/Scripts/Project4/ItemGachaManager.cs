using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.EventSystems;
using UnityEngine.UI;

public class ItemGachaManager : MonoBehaviour
{
    Perform_Gacha1 performGacha1;
    Perform_Gacha10 performGacha10;

    private void Awake()
    {
        performGacha1 = new Perform_Gacha1();
        performGacha10 = new Perform_Gacha10();
    }
    // Start is called before the first frame update
    void Start()
    {
        performGacha1.Init();
        performGacha10.Init();
    }

    // Update is called once per frame
    void Update()
    {

    }

    public void OnClickGachaButton()
    {
        int userId = UserApplication.p4_userDataManager.GetUserId();
        UserApplication.p4_phpConnectManager.CallPlayItemGacha(userId);
    }

    public void DrawPerform(string gachaResult)
    {
        string[] gachaResultArray = gachaResult.Split(",");

        performGacha1.Activate();
        FixItemManager.FixItemData fixItemData = UserApplication.fixItemManager.GetFixItemData(int.Parse(gachaResultArray[0]));
        performGacha1.SetSprite(Resources.Load<Sprite>(fixItemData.imagePath));
        performGacha1.SetName(fixItemData.name);
    }
}
