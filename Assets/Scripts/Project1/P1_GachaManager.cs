using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class P1_GachaManager : MonoBehaviour
{
    [SerializeField]
    int gachaPickCharacterId;

    Perform_Gacha1 performGacha1;

    void Awake()
    {
        gachaPickCharacterId = 0;
        performGacha1 = new Perform_Gacha1();
    }
    void Start()
    {
        performGacha1.Init();
    }

    // Update is called once per frame
    void Update()
    {
        
    }

    public void OnClickGachaButton()
    {
        gachaPickCharacterId = Random.Range(0, DefineParam.CHARA_NUM);

        Debug.Log(gachaPickCharacterId.ToString());

        DrawPerform();
    }

    public void DrawPerform()
    {
        performGacha1.Activate();
        FixCharaManager.FixCharaData fixCharaData = UserApplication.fixCharaManager.GetFixCharaData(gachaPickCharacterId);
        performGacha1.SetSprite(Resources.Load<Sprite>(fixCharaData.imagePath));
        performGacha1.SetName(fixCharaData.name);
    }

    public void OnClickGachaPerformBack()
    {
        performGacha1.Deactivate();
    }
}
