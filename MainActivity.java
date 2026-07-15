// This java package is an Android wrapper that enables the CHArT5k (Web App) to be published on Google Play Store.
// This configuration comes in four parts: (after installation of Node.js, java SDK etc.):
//      1. C:\CHArT5k-WebView\app\src\main\java\com\chart5k\app\MainActivity.java (i.e this file)
//      2. C:\CHArT5k-WebView\app\src\main\AndroidManifest.xml
//      3. C:\CHArT5k-WebView\build.gradle
//      4. C:\CHArT5k-WebView\settings.gradle
// After applying any changes, test the APK after signing...
//      a.  C:\CHArT5k-WebView\gradle-8.5\bin\gradle --stop
//          C:\CHArT5k-WebView\gradle-8.5\bin\gradle assembleRelease bundleRelease
//      b.  C:\Android-CLI\build-tools\34.0.0\apksigner.bat sign \
//              --ks C:\CHArT5k-WebView\android.keystore --ks-key-alias chart5k \
//              C:\CHArT5k-WebView\app\build\outputs\apk\release\app-release-unsigned.apk
//      c.  ren C:\CHArT5k-WebView\app\build\outputs\apk\release\app-release-unsigned.apk app-release-signed.apk
//      d.  jarsigner -verbose -sigalg SHA256withRSA -digestalg SHA256 -keystore C:\CHArT5k-WebView\android.keystore \
//              C:\CHArT5k-WebView\app\build\outputs\bundle\release\app-release.aab chart5k
// Upload .apk and .aab file for access from mobile to G Drive:
//      IT Business / Parkrun / CHArt5k-Webview

package com.chart5k.app;

import android.util.Log;
import android.os.Bundle;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.content.Intent;
import android.net.Uri;
import android.app.Activity;
import android.view.View;
import androidx.browser.customtabs.CustomTabsIntent; // for Chrome Custom Tabs for the slide-up overlay

public class MainActivity extends Activity {
    private WebView myWebView;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        
        myWebView = new WebView(this);
        WebSettings webSettings = myWebView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        webSettings.setDomStorageEnabled(true);
        webSettings.setCacheMode(WebSettings.LOAD_CACHE_ELSE_NETWORK);    // Remember your App screen changes
        // Core layout mapping options
        webSettings.setUseWideViewPort(true);
        webSettings.setLoadWithOverviewMode(true);  
        webSettings.setSupportZoom(true);           
        webSettings.setBuiltInZoomControls(true);
        webSettings.setDisplayZoomControls(false);
        myWebView.setImportantForAutofill(View.IMPORTANT_FOR_AUTOFILL_YES);    // enable DOM interaction
        
        myWebView.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, String url) {
                if (url.contains("script.google.com")) {   // Keep Web App navigation internal
                    return false; 
                } 
                if (url.contains("spreadsheets")) {        // Primary App service navigates to a single Group SS by sheet gid & range focus (
                    Log.d("CHArT5k", "Launching Spreadsheet Overlay: " + url);
                    try {     // Launch the spreadsheet as an overlay (Bubblewrap style); assume no header, toolbar etc. (rm=minimal) and NOT GSheets!
                        CustomTabsIntent.Builder builder = new CustomTabsIntent.Builder();
                        CustomTabsIntent customTabsIntent = builder.build();
                        customTabsIntent.intent.setPackage("com.android.chrome");    // -Force Browser instead of blind GSheets
                        Bundle headers = new Bundle();    // Force the custom tab to use the Desktop User-Agent for this session
                        // Standard Mobile User Agent for your core Web App (so it looks like an App)
                        String DESKTOP_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36";
                        headers.putString("User-Agent", DESKTOP_UA);    // spoof the Desktop UA for the spreadsheet overlay!
                        java.util.ArrayList<Bundle> headersList = new java.util.ArrayList<>();
                        headersList.add(headers);
                        customTabsIntent.intent.putParcelableArrayListExtra(android.provider.Browser.EXTRA_HEADERS, headersList);
                        customTabsIntent.launchUrl(MainActivity.this, Uri.parse(url));
                    } catch (Exception e) {
                        Log.d("CHArT5k", "Fallback: " + url);
                        view.loadUrl(url);    // fallback if Custom Tabs fail
                    }
                    return true; // Ensures App screen persists in the background and surfaces when returning from spreadsheet (on closure with back)
                } else {                                    // Otherwise, connect to User-preferred mailer OR default browser (htmlview only; no gid/range
                    try {
                        String actionIntent = url.startsWith("mailto:") ? Intent.ACTION_SENDTO : Intent.ACTION_VIEW;
                        Intent intent = new Intent(actionIntent, Uri.parse(url));
                        startActivity(intent);
                    } catch (Exception e) {
                        Log.d("CHArT5k", "External link: " + url);
                    }
                    return true;
                }
            }
        });
        myWebView.loadUrl("https://script.google.com/macros/s/AKfycbzGA2ARs2d8ON4xfIOKTMY5WFqE5oyNz5XLhEB_LeIzqj3mKNJdj2P84upsypi6hf96/exec");
        setContentView(myWebView);
    }
    
    @Override
    public void onBackPressed() {
        if (myWebView != null && myWebView.canGoBack()) {
            myWebView.goBack();        // Back button reverts to the persistent App WebView
        } else {        // allow for closure via "X" button on SS to effect reverting to App also
            moveTaskToBack(true);
        }
    }
}
