
namespace M8.SpreadsheetParser {
    [System.Serializable]
    public struct GoogleOAuth2InfoParse {
        public GoogleOAuth2Info installed;
    }

    [System.Serializable]
    public struct GoogleOAuth2Info {
        const string REDIRECT_URI_DEFAULT = "urn:ietf:wg:oauth:2.0:oob";

        public string client_id;
        public string project_id;
        public string auth_uri;
        public string token_uri;
        public string auth_provider_x509_cert_url;
        public string client_secret;
        public string[] redirect_uris;

        public string GetRedirectURI() {
            if(redirect_uris != null && redirect_uris.Length > 0)
                return redirect_uris[0];

            return REDIRECT_URI_DEFAULT;
        }
    };
}