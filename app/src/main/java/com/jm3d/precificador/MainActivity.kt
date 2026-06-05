package com.jm3d.precificador

import android.annotation.SuppressLint
import android.app.Activity
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.webkit.*
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.FileProvider
import androidx.core.view.WindowCompat
import com.jm3d.precificador.databinding.ActivityMainBinding
import java.io.File

class MainActivity : AppCompatActivity() {

    private lateinit var binding: ActivityMainBinding
    private var filePathCallback: ValueCallback<Array<Uri>>? = null

    private val fileChooserLauncher =
        registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
            val cb = filePathCallback ?: return@registerForActivityResult
            cb.onReceiveValue(WebChromeClient.FileChooserParams.parseResult(result.resultCode, result.data))
            filePathCallback = null
        }

    @SuppressLint("SetJavaScriptEnabled")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        WindowCompat.setDecorFitsSystemWindows(window, false)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        binding.webView.apply {
            settings.javaScriptEnabled   = true
            settings.domStorageEnabled   = true
            settings.allowFileAccess     = true
            settings.allowContentAccess  = true
            settings.builtInZoomControls = false
            settings.displayZoomControls = false
            settings.useWideViewPort     = true
            settings.loadWithOverviewMode = true

            addJavascriptInterface(Bridge(), "Bridge")

            webViewClient = object : WebViewClient() {
                override fun shouldOverrideUrlLoading(v: WebView?, r: WebResourceRequest?) = false
            }

            webChromeClient = object : WebChromeClient() {
                override fun onShowFileChooser(wv: WebView?, cb: ValueCallback<Array<Uri>>?, params: FileChooserParams?): Boolean {
                    filePathCallback?.onReceiveValue(null)
                    filePathCallback = cb
                    return try {
                        val intent = Intent(Intent.ACTION_GET_CONTENT).apply {
                            addCategory(Intent.CATEGORY_OPENABLE)
                            type = "image/*"
                        }
                        fileChooserLauncher.launch(Intent.createChooser(intent, "Selecionar imagem"))
                        true
                    } catch (e: Exception) {
                        filePathCallback = null
                        false
                    }
                }
            }
            loadUrl("file:///android_asset/index.html")
        }
    }

    inner class Bridge {
        @JavascriptInterface
        fun shareHtml(html: String, filename: String) {
            runOnUiThread {
                try {
                    val file = File(cacheDir, filename)
                    file.writeText(html, Charsets.UTF_8)
                    val uri = FileProvider.getUriForFile(this@MainActivity, "$packageName.provider", file)
                    startActivity(Intent.createChooser(
                        Intent(Intent.ACTION_SEND).apply {
                            type = "text/html"
                            putExtra(Intent.EXTRA_STREAM, uri)
                            putExtra(Intent.EXTRA_SUBJECT, "Orçamento JM 3D")
                            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                        }, "Compartilhar orçamento"
                    ))
                } catch (e: Exception) {
                    Toast.makeText(this@MainActivity, "Erro ao compartilhar.", Toast.LENGTH_SHORT).show()
                }
            }
        }

        @JavascriptInterface
        fun toast(msg: String) {
            runOnUiThread { Toast.makeText(this@MainActivity, msg, Toast.LENGTH_SHORT).show() }
        }
    }

    @Deprecated("Deprecated in Java")
    override fun onBackPressed() {
        if (binding.webView.canGoBack()) binding.webView.goBack() else super.onBackPressed()
    }
}
