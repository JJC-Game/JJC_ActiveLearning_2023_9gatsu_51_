using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.Events;
using UnityEngine.Networking;
using static P3_UserDataManager;

public class PHPConnectManager : MonoBehaviour
{
    protected string serverURL = "http://18.178.60.234/students/active_learning/";
    protected string userPHPFolderPath = "user99/sample_laravel/";
    protected string phpConnectResultText = "";
    private bool isConnecting = false;
    protected int calledUserId = 0;

    public const string ID_SAVE_KEY = "AL_2023_9GATSU_USER_ID";

    protected void CallPHPConnection(string phpFileName, UnityAction callbackFunc = null)
    {
        if (isConnecting)
        {
            Debug.LogAssertion("別の接続を実行中です.");
            return;
        }
        string url = serverURL + userPHPFolderPath + phpFileName;
        isConnecting = true;
        StartCoroutine(UrlAccess(url, () => callbackFunc(), () => CallError()));

        var request = UnityWebRequest.Get(url);
    }

    public void CallError()
    {
        Debug.Assert(false);
    }

    protected IEnumerator UrlAccess(string url, UnityAction callbackFunc = null, UnityAction errorCallbackFunc = null)
    {
        UnityWebRequest request = UnityWebRequest.Get(url);

        // リクエスト送信
        yield return request.SendWebRequest();

        if (request.result != UnityWebRequest.Result.Success)
        {
            Debug.Log("error:" + request.error);
            Debug.Log("errorURL:" + url);

            if (errorCallbackFunc != null)
            {
                errorCallbackFunc();
            }
            isConnecting = false;
            yield break;
        }
        else
        {
            // 結果をテキストとして表示します
            Debug.Log("resultText:" + request.downloadHandler.text);

            phpConnectResultText = request.downloadHandler.text;

            isConnecting = false;
            if (callbackFunc != null)
            {
                callbackFunc();
            }
        }
    }

    public void CallUserData(int userId)
    {
        calledUserId = userId;
        string phpFileName = "app_user/" + userId.ToString();
        CallPHPConnection(phpFileName, () => RefreshUserDataUI());
    }
    private void RefreshUserDataUI()
    {
        UserApplication.p3_userDataManager.SetUserData(phpConnectResultText);
    }

    public void CallPlayGacha(int userId)
    {
        calledUserId = userId;
        string phpFileName = "play_chara_gacha/" + userId.ToString();
        CallPHPConnection(phpFileName, () => RefreshGachaPerformUI());
    }
    public void CallPlayGacha10(int userId)
    {
        calledUserId = userId;
        string phpFileName = "play_chara_gacha_10/" + userId.ToString();
        CallPHPConnection(phpFileName, () => RefreshGacha10PerformUI());
    }
    private void RefreshGachaPerformUI()
    {
        GameObject obj = GameObject.Find("GachaManager");
        if (obj != null)
        {
            P3_GachaManager gachaManager = obj.GetComponent<P3_GachaManager>();
            gachaManager.DrawPerform(phpConnectResultText);
        }
        CallUserData(calledUserId);
    }

    private void RefreshGacha10PerformUI()
    {
        GameObject obj = GameObject.Find("GachaManager");
        if (obj != null)
        {
            P3_GachaManager gachaManager = obj.GetComponent<P3_GachaManager>();
            gachaManager.DrawPerform10(phpConnectResultText);
        }
        CallUserData(calledUserId);
    }

    public void CallClearHasCharaFlag(int userId)
    {
        calledUserId = userId;
        string phpFileName = "clear_has_chara_flag/" + userId.ToString();
        CallPHPConnection(phpFileName, () => CallbackClearHasCharaFlag());
    }

    private void CallbackClearHasCharaFlag()
    {
        CallUserData(calledUserId);
    }

}
