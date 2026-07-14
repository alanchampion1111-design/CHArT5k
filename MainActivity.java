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
    // potentially for external parkrun links only
    // private final String MOBILE_UA = "Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Mobile Safari/537.36";
    private final String DESKTOP_UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        // Keep alive when rotating mobile 
        if (myWebView == null) {
            myWebView = new WebView(this);
            WebSettings webSettings = myWebView.getSettings();
            webSettings.setJavaScriptEnabled(true);
            webSettings.setDomStorageEnabled(true);
            webSettings.setDatabaseEnabled(true);           // Remembers viewport states
            // Force local cache to remember your active tab/screen changes
            webSettings.setCacheMode(WebSettings.LOAD_CACHE_ELSE_NETWORK);
            // Core layout mapping options
            webSettings.setUseWideViewPort(true);       
            webSettings.setLoadWithOverviewMode(true);  
            webSettings.setSupportZoom(true);           
            webSettings.setBuiltInZoomControls(true);
            webSettings.setDisplayZoomControls(false);
            myWebView.setImportantForAutofill(View.IMPORTANT_FOR_AUTOFILL_YES);
            webSettings.setUserAgentString(DESKTOP_UA);    // essential for gid & viewports
            
            myWebView.setWebViewClient(new WebViewClient() {
                @Override
                public boolean shouldOverrideUrlLoading(WebView view, String url) {
                    if (url.startsWith("mailto:")) {
                        try {       // handle mail extrrnally via user's preferred emailer
                            Intent intent = new Intent(Intent.ACTION_SENDTO, Uri.parse(url));
                            startActivity(intent);
                        } catch (Exception e) {}
                        return true;
                    } else {        // retain same UA if internal to App or link to Group SS
                        if (url.contains("script.google.com") || url.contains("docs.google.com"))
                            return false;   
                        else {      // otherwise assume external parkrun link (via browser
                            try {
                                Intent intent = new Intent(Intent.ACTION_VIEW, Uri.parse(url));
                                startActivity(intent);
                            } catch (Exception e) {}
                                return true;    // handle link externally; (like with mailto:)
                            }
                        }
                    }
                }
            });
            
            if (savedInstanceState != null) {
                myWebView.restoreState(savedInstanceState);
            } else {
                myWebView.loadUrl("https://script.google.com/macros/s/AKfycbzGA2ARs2d8ON4xfIOKTMY5WFqE5oyNz5XLhEB_LeIzqj3mKNJdj2P84upsypi6hf96/exec");
            }
        }
        setContentView(myWebView);
    }
    
    @Override
    protected void onSaveInstanceState(Bundle outState) {
        super.onSaveInstanceState(outState);
        if (myWebView != null) {
            myWebView.saveState(outState);
        }
    }

    @Override
    protected void onRestoreInstanceState(Bundle savedInstanceState) {
        super.onRestoreInstanceState(savedInstanceState);
        if (myWebView != null) {
            myWebView.restoreState(savedInstanceState);
        }
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
