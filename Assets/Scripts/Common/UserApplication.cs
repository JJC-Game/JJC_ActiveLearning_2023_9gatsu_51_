using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class UserApplication : MonoBehaviour
{
    private void Awake()
    {
        S = this;
    }
    // Start is called before the first frame update
    void Start()
    {

    }

    static private UserApplication _S;
    static private UserApplication S
    {
        get
        {
            if (_S == null)
            {
                return null;
            }
            return _S;
        }
        set
        {
            if (_S != null)
            {
                Debug.LogError("_Sは既に設定されています.");
            }
            _S = value;
        }
    }

    [Header("UserApplicationからアクセスできるコンポーネント")]
    public FixCharaManager _fixCharaManager;
    static public FixCharaManager fixCharaManager
    {
        get
        {
            if (S._fixCharaManager == null)
            {
                S._fixCharaManager = GameObject.Find("Common").GetComponent<FixCharaManager>();
            }
            return S._fixCharaManager;
        }
    }
    
    public UserDataManager _userDataManager;
    static public UserDataManager userDataManager
    {
        get
        {
            if (S._userDataManager == null)
            {
                S._userDataManager = GameObject.Find("UserDataManager").GetComponent<UserDataManager>();
            }
            return S._userDataManager;
        }
    }

    public CharaGridRenderer _charaGridRenderer;
    static public CharaGridRenderer charaGridRenderer
    {
        get
        {
            if (S._charaGridRenderer == null)
            {
                S._charaGridRenderer = GameObject.Find("UIParts_CharaGridRenderer").GetComponent<CharaGridRenderer>();
            }
            return S._charaGridRenderer;
        }
    }

}
