using System.Collections;
using System.Collections.Generic;
using UnityEngine;

namespace M8.SpreadsheetParser {
    [CreateAssetMenu(menuName = "M8/Spreadsheet Parser/Google Settings")]
    public class GoogleSettings : ScriptableObject {
        [HideInInspector]
        public string jsonFilePath = string.Empty;

        public GoogleOAuth2Info authInfo;

        [Header("Account (if required)")]
        public string username;
        public string password;

        // enter Access Code after getting it from auth url
        [HideInInspector]
        public string accessCode = "";

        // enter Auth 2.0 Refresh Token and AccessToken after succesfully authorizing with Access Code
        [HideInInspector]
        public string refreshToken = "";

        [HideInInspector]
        public string accessToken = "";
    }
}