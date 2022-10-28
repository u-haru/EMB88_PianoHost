#include <Windows.h>
#include <conio.h>
#include <stdio.h>
#include <string.h>
// status
#define NOTE_OFF 0x8
#define NOTE_ON 0x9
#define MODULATION 0x1
#define RELEASETIME 0x48
#define CONTROL_CHANGE 0xb
#define PROGRAM_CHANGE 0xc

#define MAKESMSG(st, ch, d1, d2) ((st) << 4 | (ch) | (d1) << 8 | (d2) << 16)
#define MAKESMSG_(d1, d2) ((d1) | (d2) << 8)

// http://titech-ssr.blog.jp/archives/1048225188.html
HANDLE comInit(char* portName, int baudrate) {
	HANDLE hComPort = CreateFile(  //ファイルとしてポートを開く
	    portName,  // ポート名を指すバッファへのポインタ:COM4を開く（デバイスマネージャでどのポートが使えるか確認）
	    GENERIC_READ | GENERIC_WRITE,  // アクセスモード:読み書き両方する
	    0,  //ポートの共有方法を指定:オブジェクトは共有しない
	    NULL,  //セキュリティ属性:ハンドルを子プロセスへ継承しない
	    OPEN_EXISTING,  //ポートを開き方を指定:既存のポートを開く
	    0,  //ポートの属性を指定:同期非同期にしたいときはFILE_FLAG_OVERLAPPED
	    NULL  // テンプレートファイルへのハンドル:NULLって書け
	);
	if (hComPort == INVALID_HANDLE_VALUE) {  //ポートの取得に失敗
		printf("指定COMポートが開けません.\n");
		CloseHandle(hComPort);  //ポートを閉じる
		return NULL;
	} else {
		printf("COMポートは正常に開けました.\n");
	}

	if (SetupComm(hComPort,  // COMポートのハンドラ
	              1024,      //受信バッファサイズ:1024byte
	              1024       //送信バッファ:1024byte
	              )) {
		printf("送受信バッファの設定が完了しました.\r\n");
	} else {
		printf("送受信バッファの設定ができません.\r\n");
		CloseHandle(hComPort);
		return NULL;
	}
	if (PurgeComm(hComPort,  // COMポートのハンドラ
	              PURGE_TXABORT | PURGE_RXABORT | PURGE_TXCLEAR |
	                  PURGE_RXCLEAR  //出入力バッファをすべてクリア
	              )) {
		printf("送受信バッファの初期化が完了しました.\r\n");
	} else {
		printf("送受信バッファの初期化ができません.\r\n");
		CloseHandle(hComPort);
		return NULL;
	}
	DCB dcb;                       //構成情報を記録する構造体の生成
	GetCommState(hComPort, &dcb);  //現在の設定値を読み込み
	dcb.DCBlength = sizeof(DCB);   // DCBのサイズ
	dcb.BaudRate = baudrate;          //ボーレート:9600bps
	dcb.ByteSize = 8;              //データサイズ:8bit
	dcb.fBinary = TRUE;            //バイナリモード:通常TRUE
	dcb.fParity = NOPARITY;  //パリティビット:パリティビットなし
	dcb.StopBits = ONESTOPBIT;  //ストップビット:1bit
	dcb.fOutxCtsFlow = FALSE;   // CTSフロー制御:フロー制御なし
	dcb.fOutxDsrFlow = FALSE;  // DSRハードウェアフロー制御：使用しない
	dcb.fDtrControl = DTR_CONTROL_DISABLE;  // DTR有効/無効:DTR無効
	dcb.fRtsControl = RTS_CONTROL_DISABLE;  // RTSフロー制御:RTS制御なし

	dcb.fOutX = FALSE;  //送信時XON/XOFF制御の有無:なし
	dcb.fInX = FALSE;   //受信時XON/XOFF制御の有無:なし
	dcb.fTXContinueOnXoff =
	    TRUE;  // 受信バッファー満杯＆XOFF受信後の継続送信可否:送信可
	dcb.XonLim = 512;  // XONが送られるまでに格納できる最小バイト数:512
	dcb.XoffLim = 512;  // XOFFが送られるまでに格納できる最小バイト数:512
	dcb.XonChar = 0x11;  //送信時XON文字 ( 送信可：ビジィ解除 )
	                     //の指定:XON文字として11H ( デバイス制御１：DC1 )
	dcb.XoffChar =
	    0x13;  // XOFF文字（送信不可：ビジー通告）の指定:XOFF文字として13H
	           // ( デバイス制御3：DC3 )

	dcb.fNull = TRUE;          // NULLバイトの破棄:破棄する
	dcb.fAbortOnError = TRUE;  //エラー時の読み書き操作終了:終了する
	dcb.fErrorChar =
	    FALSE;  // パリティエラー発生時のキャラクタ（ErrorChar）置換:なし
	dcb.ErrorChar = 0x00;  // パリティエラー発生時の置換キャラクタ
	dcb.EofChar =
	    0x03;  // データ終了通知キャラクタ:一般に0x03(ETX)がよく使われます。
	dcb.EvtChar = 0x02;  // イベント通知キャラクタ:一般に0x02(STX)がよく使われます

	if (SetCommState(hComPort, &dcb)) {  //エラーチェック
		printf("COMポート構成情報を変更しました.\r\n");
	} else {
		printf("COMポート構成情報の変更に失敗しました.\r\n");
		CloseHandle(hComPort);
		return NULL;
	}

	COMMTIMEOUTS TimeOut;  // COMMTIMEOUTS構造体の変数を宣言
	GetCommTimeouts(hComPort, &TimeOut);  // タイムアウトの設定状態を取得

	TimeOut.ReadTotalTimeoutMultiplier =
	    0;  //読込の１文字あたりの時間:タイムアウトなし
	TimeOut.ReadTotalTimeoutConstant = 1000;  //読込エラー検出用のタイムアウト時間
	//(受信トータルタイムアウト) = ReadTotalTimeoutMultiplier × (受信予定バイト数)
	//+ ReadTotalTimeoutConstant
	TimeOut.WriteTotalTimeoutMultiplier =
	    0;  //書き込み１文字あたりの待ち時間:タイムアウトなし
	TimeOut.WriteTotalTimeoutConstant =
	    1000;  //書き込みエラー検出用のタイムアウト時間
	//(送信トータルタイムアウト) = WriteTotalTimeoutMultiplier ×(送信予定バイト数)
	//+ WriteTotalTimeoutConstant

	if (SetCommTimeouts(hComPort, &TimeOut)) {  //エラーチェック
		printf("タイムアウトの設定に成功しました.\r\n");
	} else {
		printf("タイムアウトの設定に失敗しました.\r\n");
		CloseHandle(hComPort);
		return NULL;
	}
	return hComPort;
}

int main(int argc, char* argv[]) {
	if (argc < 2) {
		printf("Usage: %s comport\n", argv[0]);
		return -1;
	}
	HANDLE com = comInit(argv[1],38400);

    if(com==NULL){
        printf("err");
        return -1;
    }
	int C4 = 4;
	bool beforeState[29] = {0};
	u_int pitch = 1;
    const char keyb[25] = {(char)VK_LSHIFT,'A','Z','S','X','D','C','F','V','G','B','H','N','J','M','K',(char)VK_OEM_COMMA,'L',(char)VK_OEM_PERIOD,(char)VK_OEM_PLUS,(char)VK_OEM_2,(char)VK_OEM_1,(char)VK_OEM_102,(char)VK_OEM_6,(char)VK_RSHIFT};
    bool noisestate[9] = {0};
    bool dtmfstate[12] = {0};
    char note = 0;
	u_int speckey = sizeof(keyb) / sizeof(const char);
	memset(beforeState, 0, sizeof(beforeState));

	char buf = 0;
	while (!GetAsyncKeyState(VK_ESCAPE)) {
		for (int i = 0; i < 25; i++) {
			note = i;
			if (i - C4 >= 0)
				for (int j = 0; j <= (i / 14); j++) {
					if (i - C4 - (j * 14) > 5) note--;
					if (i - C4 - (j * 14) > 13) note--;
				}
			else
				for (int j = 0; j <= (i / 14); j++) {
					if ((j * 14) - i + C4> 9) note++;
					if ((j * 14) - i + C4> 1) note++;
				}
			buf = (C4-3 + note + pitch * 12);
			if (GetAsyncKeyState(keyb[i] & 0xff) && !beforeState[i]) {
				beforeState[i] = TRUE; 
                printf("On\tdata:%.2x\t%d\n",buf,buf&0x7F);
                buf |= 0x80;
				if ((i-C4)%14!=5&&(i-C4)%14!=13&&(i-C4)%14!=-9&&(i-C4)%14!=-1)
					WriteFile(com, &buf, sizeof(buf), NULL, NULL);
			}
			if (!GetAsyncKeyState(keyb[i] & 0xff) && beforeState[i]&&!GetAsyncKeyState(VK_SPACE)) {
			    WriteFile(com, &buf, sizeof(buf), NULL, NULL);
				beforeState[i] = FALSE;
                printf("Off\tdata:%x\t %d\n",buf,buf);}
		}

		if (GetAsyncKeyState('Q'))pitch=0;
		if (GetAsyncKeyState('W'))pitch=1;
		if (GetAsyncKeyState('E'))pitch=2;
		if (GetAsyncKeyState('R'))pitch=3;

		for(int i = 0 ;i < 9;i++){
            buf = 55+i;
			if (GetAsyncKeyState(i+49) && !noisestate[i]) {
				noisestate[i] = TRUE; 
                printf("On\tdata:%.2x\t%d\n",buf,buf&0x7F);
                buf |= 0x80;
				WriteFile(com, &buf, sizeof(buf), NULL, NULL);
			}
			if (!GetAsyncKeyState(i+49) && noisestate[i]&&!GetAsyncKeyState(VK_SPACE)) {
			    WriteFile(com, &buf, sizeof(buf), NULL, NULL);
				noisestate[i] = FALSE;
                printf("Off\tdata:%x\t %d\n",buf,buf);
            }
        }
        
		for(int i = 0 ;i < 12;i++){
			if (GetAsyncKeyState(i+96) && !dtmfstate[i]) {
				dtmfstate[i] = TRUE;
				switch(i){
					case 1:case 2:case 3: buf=30;break;
					case 4:case 5:case 6: buf=32;break;
                    case 7:case 8:case 9: buf=33;break;
                    case 10:case 0:case 11: buf=35;break;
				}
                printf("On\tdata:%.2x\t%d\n",buf,buf&0x7F);
                buf |= 0x80;
				WriteFile(com, &buf, sizeof(buf), NULL, NULL);
				switch(i){
					case 1:case 4:case 7:case 10: buf=40;break;
					case 2:case 5:case 8:case 0: buf=41;break;
					case 3:case 6:case 9:case 11: buf=43;break;
				}
                printf("On\tdata:%.2x\t%d\n",buf,buf&0x7F);
                buf |= 0x80;
				WriteFile(com, &buf, sizeof(buf), NULL, NULL);
            }
			if (!GetAsyncKeyState(i+96) && dtmfstate[i]&&!GetAsyncKeyState(VK_SPACE)) {
				dtmfstate[i] = FALSE;
				switch(i){
					case 1:case 2:case 3: buf=30;break;
					case 4:case 5:case 6: buf=32;break;
                    case 7:case 8:case 9: buf=33;break;
                    case 10:case 0:case 11: buf=35;break;
				}
                printf("Off\tdata:%.2x\t%d\n",buf,buf&0x7F);
				WriteFile(com, &buf, sizeof(buf), NULL, NULL);
				switch(i){
					case 1:case 4:case 7:case 10: buf=40;break;
					case 2:case 5:case 8:case 0: buf=41;break;
					case 3:case 6:case 9:case 11: buf=43;break;
				}
                printf("Off\tdata:%.2x\t%d\n",buf,buf&0x7F);
				WriteFile(com, &buf, sizeof(buf), NULL, NULL);
            }
        }
	}
    return 0;
}