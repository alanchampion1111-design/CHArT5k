// This java package is an Android wrapper that enables the CHArT5k (Wb App) to be published on Google Play Store.
// This configuration comes in two parts: (after installation of Node.js, java SDK etc.)
//      1. C:\CHArT5k-WebView\app\src\main\java\com\chart5k\app\MainActivity.java (i.e this file)
//      2. C:\CHArT5k-WebView\app\src\main\AndroidManifest.xml
// After applying any changes, test the APK after signing
//      a.  C:\CHArT5k-WebView\gradle-8.5\bin\gradle --stop
//          C:\CHArT5k-WebView\gradle-8.5\bin\gradle assembleRelease bundleRelease
//      b.  C:\Android-CLI\build-tools\34.0.0\apksigner.bat sign \
//              --ks C:\CHArT5k-WebView\android.keystore --ks-key-alias chart5k \
//              C:\CHArT5k-WebView\app\build\outputs\apk\release\app-release-unsigned.apk
//      c.  del C:\CHArT5k-WebView\app\build\outputs\apk\release\app-release-signed.apk
//          ren C:\CHArT5k-WebView\app\build\outputs\apk\release\app-release-unsigned.apk app-release-signed.apk
// Upload apk file for access from mobile to G Drive:
//      IT Business / Parkrun / CHArt5k-Webview

package com.chart5k.app;

import android.os.Bundle;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.content.Intent;
import android.net.Uri;
import android.app.Activity;
import android.view.View;

public class MainActivity extends Activity {
    private WebView myWebView;
    private final String MOBILE_UA = "Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Mobile Safari/537.36";
    private final String DESKTOP_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        // Keep a single instance on orientation change
        if (myWebView == null) {
            myWebView = new WebView(this);
            WebSettings webSettings = myWebView.getSettings();
            webSettings.setJavaScriptEnabled(true);
            webSettings.setDomStorageEnabled(true);
            webSettings.setDatabaseEnabled(true);           // Remembers viewport states
            // Tells WebView to pull from local cache on Back navigation to keep your exact tab screen
            webSettings.setCacheMode(WebSettings.LOAD_CACHE_ELSE_NETWORK);
            // Core layout mapping options
            webSettings.setUseWideViewPort(true);       
            webSettings.setLoadWithOverviewMode(true);  
            webSettings.setSupportZoom(true);           
            webSettings.setBuiltInZoomControls(true);   
            webSettings.setDisplayZoomControls(false);
            
            myWebView.setImportantForAutofill(View.IMPORTANT_FOR_AUTOFILL_YES);
            webSettings.setUserAgentString(MOBILE_UA);
            myWebView.setWebViewClient(new WebViewClient() {
                @Override
                public boolean shouldOverrideUrlLoading(WebView view, String url) {
                    if (url.startsWith("mailto:")) {
                        try {
                            Intent intent = new Intent(Intent.ACTION_SENDTO, Uri.parse(url));
                            startActivity(intent);
                        } catch (Exception e) {}
                        return true;
                    }
                    
                    if (url.contains("docs.google.com/spreadsheets")) {
                        // Fix 3: Try to open in the native Google Sheets App if installed
                        try {
                            Intent intent = new Intent(Intent.ACTION_VIEW, Uri.parse(url));
                            intent.setPackage("com.google.android.apps.docs.editors.sheets");
                            intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
                            view.getContext().startActivity(intent);
                            return true; // Successfully handed off to native app
                        } catch (Exception e) {
                            // Fallback if Google Sheets App is not installed: load in WebView as Desktop
                            view.stopLoading(); // Fix 1: Stop any active loading to prevent random page refresh
                            view.getSettings().setUserAgentString(DESKTOP_UA);
                        }
                    } else {
                        view.getSettings().setUserAgentString(MOBILE_UA);
                    }

                    return false; 
                }
            });
            // native link to Goohgle Web App
            myWebView.loadUrl("https://script.google.com/macros/s/AKfycbzGA2ARs2d8ON4xfIOKTMY5WFqE5oyNz5XLhEB_LeIzqj3mKNJdj2P84upsypi6hf96/exec");
        }
        setContentView(myWebView);
    }

    @Override
    public void onBackPressed() {
        if (myWebView != null && myWebView.canGoBack()) {
            myWebView.goBack(); // Navigates backward historically step-by-step
        } else {
            super.onBackPressed(); 
        }
    }
}
