package com.bedroomcomputing.googlespreadsheetapi

import android.content.Intent
import android.os.Bundle
import android.util.Log
import android.widget.Button
import androidx.appcompat.app.AppCompatActivity
import com.google.android.gms.auth.api.signin.GoogleSignIn
import com.google.android.gms.auth.api.signin.GoogleSignInAccount
import com.google.android.gms.auth.api.signin.GoogleSignInClient
import com.google.android.gms.auth.api.signin.GoogleSignInOptions
import com.google.android.gms.common.SignInButton
import com.google.android.gms.common.api.ApiException
import com.google.android.gms.common.api.Scope
import com.google.android.gms.tasks.Task
import com.google.api.client.extensions.android.http.AndroidHttp
import com.google.api.client.googleapis.extensions.android.gms.auth.GoogleAccountCredential
import com.google.api.client.json.jackson2.JacksonFactory
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.model.AppendValuesResponse
import com.google.api.services.sheets.v4.model.ValueRange
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.MainScope
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.util.*


class MainActivity : AppCompatActivity() {

    lateinit var mGoogleSignInClient: GoogleSignInClient
    lateinit var  credential: GoogleAccountCredential

    // サインイン用intentを識別するためのID。0であることに意味はない
    val RC_SIGN_IN = 0

    // SheetServiceの初期化
    val transport = AndroidHttp.newCompatibleTransport();
    val factory = JacksonFactory.getDefaultInstance()

    override fun onCreate(savedInstanceState: Bundle?) {

        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        // サインインのオプションを設定。Emailの取得とspreadsheetのアクセスを要求する
        val gso= GoogleSignInOptions.Builder(GoogleSignInOptions.DEFAULT_SIGN_IN)
                .requestScopes(Scope("https://www.googleapis.com/auth/spreadsheets"))
                .requestEmail()
                .build()

        mGoogleSignInClient = GoogleSignIn.getClient(this, gso)
        credential = GoogleAccountCredential.usingOAuth2(this, Collections.singleton("https://www.googleapis.com/auth/spreadsheets"))

        // もし前回起動時にサインインしていたら、サインイン不要
        val account = GoogleSignIn.getLastSignedInAccount(this)
        account?.let{
            Log.i("Main", "${account.displayName}")
            credential?.setSelectedAccount(account.account)
        }

        // サインイン
        findViewById<SignInButton>(R.id.sign_in_button).setOnClickListener{
            signIn()
        }

        // サインアウト
        findViewById<Button>(R.id.button_signout).setOnClickListener{
            signOut()
        }

        // 読み込み。ネットワーク通信なので、coroutine内で行う。
        findViewById<Button>(R.id.button_read).setOnClickListener{
            MainScope().launch{
                withContext(Dispatchers.Default){
                    read()
                }
            }
        }
        // 書き込み。ネットワーク通信なので、coroutine内で行う。
        findViewById<Button>(R.id.button_write).setOnClickListener{
            MainScope().launch{
                withContext(Dispatchers.Default){
                    write()
                }
            }
        }
    }

    private fun signIn() {
        // サインイン用のインテントを呼び出す。onActivityResultに戻ってくる
        val signInIntent: Intent = mGoogleSignInClient.getSignInIntent()
        startActivityForResult(signInIntent, RC_SIGN_IN)
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)

        // サインイン完了時の処理
        if (requestCode == RC_SIGN_IN) {
            val task = GoogleSignIn.getSignedInAccountFromIntent(data)
            handleSignInResult(task)
        }
    }

    private fun handleSignInResult(completedTask: Task<GoogleSignInAccount>) {
        try {
            val account = completedTask.getResult(ApiException::class.java)

            Log.i("Main", "${account?.displayName}")
            credential?.setSelectedAccount(account?.account)

        } catch (e: ApiException) {
            Log.w("Main", "signInResult:failed code=" + e.statusCode)

        }
    }

    fun read(){
        val sheetsService = Sheets.Builder(AndroidHttp.newCompatibleTransport(), JacksonFactory.getDefaultInstance(), credential)
                // アプリケーション名を指定するのだが、適当でいいっぽい
                .setApplicationName("Test Project2")
                .build();

        // 値取得
        val response = sheetsService
                .spreadsheets().values()
                // Spreadsheet idはURLのhttps://docs.google.com/spreadsheets/d/xxxx/...のxxx部分
                .get("1XoRcqhbAkhYB8zh_k4_VTh3_V6TrJLeTz9NfW5mTY_8", "Sheet1!A1")
                .execute()

        val values = response.getValues()

        // Stringにキャストする
        val a1 =  values[0][0] as String

        Log.i("MainActivity", a1)
    }

     fun write(){
        val sheetsService = Sheets.Builder(AndroidHttp.newCompatibleTransport(), JacksonFactory.getDefaultInstance(), credential)
                // アプリケーション名を指定するのだが、適当でいいっぽい
                .setApplicationName("Test Project2")
                .build();

         // 二次元配列で書き込む値を保持
        val values: List<List<Any>> = Arrays.asList(
                Arrays.asList("A","B") ,
                Arrays.asList("C","D") // Additional rows ...
        )
        val body = ValueRange().setValues(values)

         // 書き込み。
         val result: AppendValuesResponse =
                sheetsService.spreadsheets().values()
                        // appendは一番下の行に追加していってくれる
                        .append("1XoRcqhbAkhYB8zh_k4_VTh3_V6TrJLeTz9NfW5mTY_8", "Sheet1!A:B", body)
                        // RAWを指定すると値がそのまま表示される。USER_ENTEREDだと数字や日付の書式が手入力時と同じになるらしい
                        .setValueInputOption("RAW")
                        .execute()

    }

    private fun signOut() {
        mGoogleSignInClient.signOut()
                .addOnCompleteListener(this) {
                }
    }

}