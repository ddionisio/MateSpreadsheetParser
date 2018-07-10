using UnityEngine;
using UnityEditor;
using System;
using System.Collections.Generic;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace M8.SpreadsheetParser {
    public class GoogleRequest {
        const string SCOPE = "https://www.googleapis.com/auth/drive https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
        const string TOKEN_TYPE = "refresh";

        public static GOAuth2RequestFactory RefreshAuthenticate(GoogleSettings googleSettings) {
            if(string.IsNullOrEmpty(googleSettings.refreshToken) 
                || string.IsNullOrEmpty(googleSettings.accessToken) 
                || string.IsNullOrEmpty(googleSettings.authInfo.client_id))
                return null;

            OAuth2Parameters parameters = new OAuth2Parameters() {
                RefreshToken = googleSettings.refreshToken,
                AccessToken = googleSettings.accessToken,
                ClientId = googleSettings.authInfo.client_id,
                ClientSecret = googleSettings.authInfo.client_secret,
                Scope = "https://www.googleapis.com/auth/drive https://spreadsheets.google.com/feeds",
                AccessType = "offline",
                TokenType = "refresh"
            };
            return new GOAuth2RequestFactory("spreadsheet", "MySpreadsheetIntegration-v1", parameters);
        }

        public static void InitAuthenticate(GoogleSettings googleSettings) {
            string clientId = googleSettings.authInfo.client_id;
            string clientSecret = googleSettings.authInfo.client_secret;
            string accessCode = googleSettings.accessCode;

            // OAuth2Parameters holds all the parameters related to OAuth 2.0.
            OAuth2Parameters parameters = new OAuth2Parameters();
            parameters.ClientId = clientId;
            parameters.ClientSecret = clientSecret;
            parameters.RedirectUri = googleSettings.authInfo.GetRedirectURI();

            // Retrieves the Authorization URL
            parameters.Scope = SCOPE;
            parameters.AccessType = "offline"; // IMPORTANT 
            parameters.TokenType = TOKEN_TYPE; // IMPORTANT 

            string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
            Debug.Log(authorizationUrl);
            //Debug.Log("Please visit the URL above to authorize your OAuth "
            //      + "request token.  Once that is complete, type in your access code to "
            //      + "continue...");

            parameters.AccessCode = accessCode;

            if(IsValidURL(authorizationUrl)) {
                Application.OpenURL(authorizationUrl);
                const string message = @"Copy the 'Access Code' on your browser into the access code textfield.";
                EditorUtility.DisplayDialog("Info", message, "OK");
            }
            else
                EditorUtility.DisplayDialog("Error", "Invalid URL: " + authorizationUrl, "OK");
        }

        /// <summary>
        ///  Check whether the given string is a valid http or https URL.
        /// </summary>
        private static bool IsValidURL(string url) {
            Uri uriResult;
            return (Uri.TryCreate(url, UriKind.Absolute, out uriResult) &&
                                (uriResult.Scheme == Uri.UriSchemeHttp ||
                                 uriResult.Scheme == Uri.UriSchemeHttps));
        }

        public static void FinishAuthenticate(GoogleSettings googleSettings) {
            try {
                OAuth2Parameters parameters = new OAuth2Parameters();
                parameters.ClientId = googleSettings.authInfo.client_id;
                parameters.ClientSecret = googleSettings.authInfo.client_secret;
                parameters.RedirectUri = googleSettings.authInfo.GetRedirectURI();

                parameters.Scope = SCOPE;
                parameters.AccessType = "offline"; // IMPORTANT 
                parameters.TokenType = TOKEN_TYPE; // IMPORTANT 

                parameters.AccessCode = googleSettings.accessCode;

                OAuthUtil.GetAccessToken(parameters);
                string accessToken = parameters.AccessToken;
                string refreshToken = parameters.RefreshToken;
                //Debug.Log("OAuth Access Token: " + accessToken + "\n");
                //Debug.Log("OAuth Refresh Token: " + refreshToken + "\n");

                googleSettings.refreshToken = refreshToken;
                googleSettings.accessToken = accessToken;
            }
            catch(Exception e) {
                // To display the error message with EditorGUI.Dialogue, we throw it again.
                throw new Exception(e.Message);
            }
        }
    }
}