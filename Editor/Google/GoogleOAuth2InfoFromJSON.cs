using UnityEngine;

namespace M8.SpreadsheetParser {
    [System.Serializable]
    public struct GoogleOAuth2InfoFromJSON {
        public GoogleOAuth2Info installed;

        public static GoogleOAuth2Info Parse(string json) {
            var parse = JsonUtility.FromJson<GoogleOAuth2InfoFromJSON>(json);
            return parse.installed;
        }
    }
}