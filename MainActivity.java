// This java package is an Android wrapper that enables the CHArT5k (Wb App) to be published on Google Play Store.
// This configuration comes in two parts: (after installation of Node.js, java SDK etc.)
//      1. C:\CHArT5k-WebView\app\src\main\java\com\chart5k\app\MainActivity.java (i.e this file)
//      2. C:\CHArT5k-WebView\app\src\main\AndroidManifest.xml
// After applying any changes, test the APK after signing
//      a.  C:\CHArT5k-WebView\gradle-8.5\bin\gradle --stop
//          C:\CHArT5k-WebView\gradle-8.5\bin\gradle assembleRelease bundleRelease
//      b.  C:\Android-CLI\build-tools\34.0.0\apksigner.bat sign \
//              --ks C:\CHArT5k-WebView\android.keystore --ks-key-alias chart5k \
//      c.  del C:\CHArT5k-WebView\app\build\outputs\apk\release\app-release-signed.apk
//          ren C:\CHArT5k-WebView\app\build\outputs\apk\release\app-release-unsigned.apk app-release-signed.apk
//      d.  C:\CHArT5k-WebView\app\build\outputs\apk\release\app-release-unsigned.apk \
//              jarsigner -verbose -sigalg SHA256withRSA -digestalg SHA256 -keystore \
//              C:\CHArT5k-WebView\android.keystore \
//              C:\CHArT5k-WebView\app\build\outputs\bundle\release\app-release-unsigned.aab chart5k
//      e.  del C:\CHArT5k-WebView\app\build\outputs\bundle\release\app-release-signed.aab
//          ren C:\CHArT5k-WebView\app\build\outputs\bundle\release\app-release-unsigned.aab app-release-signed.aab
// Upload .apk and .aab file for access from mobile to G Drive:
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

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        myWebView = new WebView(this);
        WebSettings webSettings = myWebView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        webSettings.setDomStorageEnabled(true);
        // Force local cache to remember your active tab/screen changes
        webSettings.setCacheMode(WebSettings.LOAD_CACHE_ELSE_NETWORK);
        // Core layout mapping options
        webSettings.setUseWideViewPort(true);       
        webSettings.setLoadWithOverviewMode(true);  
        webSettings.setSupportZoom(true);           
        webSettings.setBuiltInZoomControls(true);
        webSettings.setDisplayZoomControls(false);
        myWebView.setImportantForAutofill(View.IMPORTANT_FOR_AUTOFILL_YES);
        // webSettings.setUserAgentString(DESKTOP_UA); 
        
        myWebView.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, String url) {
                if (url.contains("script.google.com")) {
                    return false;    // action internal to App
                }    // otherwise for mail requests, for links to Group SS, or external links to parkrun 
                String actionIntent = url.startsWith("mailto:")
                    ? Intent.ACTION_SENDTO
                    : Intent.ACTION_VIEW;// assume  (url.startsWith("mailto:") || url.contains("docs.google.com"))
                try {       // handle externally via user's preferred emailer or browser
                    Intent intent = new Intent(actionIntent, Uri.parse(url));
                    startActivity(intent);
                } catch (Exception e) {}
                return true;
            }
        });
        
        myWebView.loadUrl("https://script.google.com/macros/s/AKfycbzGA2ARs2d8ON4xfIOKTMY5WFqE5oyNz5XLhEB_LeIzqj3mKNJdj2P84upsypi6hf96/exec");
        setContentView(myWebView);
    }
    
    @Override
    public void onBackPressed() {
        if (myWebView != null && myWebView.canGoBack()) {
            myWebView.goBack(); // Navigates backward historically step-by-step
        } else {
            moveTaskToBack(true);
        }
    }
}
