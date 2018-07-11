using UnityEngine;
using UnityEditor;
using System.Collections;
using System.Collections.Generic;
using System;
using System.IO;
using System.Text;

using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace M8.SpreadsheetParser {
    [CustomEditor(typeof(GoogleSettings))]
    public class GoogleSettingsInspector : Editor {
        public static bool Validator(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors policyErrors) {
            //Debug.Log("Validation successful!");
            return true;
        }

        void OnEnable() {
            // resolve TlsException error
            ServicePointManager.ServerCertificateValidationCallback += Validator;
        }

        void OnDisable() {
            // resolve TlsException error
            ServicePointManager.ServerCertificateValidationCallback -= Validator;
        }

        public override void OnInspectorGUI() {
            var dat = target as GoogleSettings;
                                    
            //other settings
            base.OnInspectorGUI();

            //file select
            EditorGUILayout.Separator();
                        
            GUILayout.BeginHorizontal(); // Begin json file setting
            GUILayout.Label("JSON File:", GUILayout.Width(90f));

            string path = "";
            if(string.IsNullOrEmpty(dat.jsonFilePath))
                path = Application.dataPath;
            else
                path = dat.jsonFilePath;

            path = EditorGUILayout.TextField(path, GUILayout.Width(250));
            if(GUILayout.Button("...", GUILayout.Width(25f))) {
                string folder = Path.GetDirectoryName(path);
                path = EditorUtility.OpenFilePanel("Open JSON file", folder, "json");
            }
            GUILayout.EndHorizontal(); // End json file setting.

            //update jsonFilePath
            if(dat.jsonFilePath != path) {
                Undo.RecordObject(dat, "Changed JSON File path.");
                dat.jsonFilePath = path;
            }

            bool lastEnabled = GUI.enabled;
            GUI.enabled = path.GetExtension() == "json" && File.Exists(path);

            //fill data from JSON file
            if(GUILayout.Button("Refresh Data From JSON")) {
                var json = File.ReadAllText(path);

                dat.authInfo = GoogleOAuth2InfoFromJSON.Parse(json);

                EditorUtility.SetDirty(dat);
            }

            GUI.enabled = lastEnabled;

            //functions
            EditorGUILayout.Separator();

            GUILayout.BeginVertical(GUI.skin.box);

            if(GUILayout.Button(new GUIContent("Start Authentication", "Click this to get an Access Code."))) {
                GoogleRequest.InitAuthenticate(dat);
            }

            var accessCode = EditorGUILayout.TextField(new GUIContent("Access Code", "Paste code generated from the page launched via Start Authentication."), dat.accessCode);
            if(dat.accessCode != accessCode) {
                Undo.RecordObject(dat, "Set Access Code");
                dat.accessCode = accessCode;
            }

            lastEnabled = GUI.enabled;
            GUI.enabled = !string.IsNullOrEmpty(dat.accessCode);

            if(GUILayout.Button("Finish Authentication")) {
                try {
                    GoogleRequest.FinishAuthenticate(dat);
                    EditorUtility.SetDirty(dat);
                }
                catch(Exception e) {
                    EditorUtility.DisplayDialog("Error", e.Message, "OK");
                }
            }

            GUI.enabled = lastEnabled;

            EditorGUILayout.LabelField("Refresh Token", dat.refreshToken);
            EditorGUILayout.LabelField("Access Token", dat.accessToken);

            GUILayout.EndVertical();
        }
    }
}