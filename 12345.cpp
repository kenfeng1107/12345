// 網路設定 const char* ssid = "Mi 10T Pro"; // WiFi 名稱 const char* password = "ohshit1234"; // WiFi 密碼 
// 取得網路時間相關參數設定 const char* ntpServer = "time.google.com"; // NTP 時間伺服器 
const long gmtOffset_sec = 28800; // 台灣時區偏移量 (8小時) 
const int daylightOffset_sec = 0; // 日光節約時間偏移 (台灣無日光節約時間) 
// 七段顯示器腳位設定 
#define pin_DIG1 22 // 第一位數的控制腳位 
#define pin_DIG2 21 // 第二位數的控制腳位 
#define pin_DIG3 2 // 第三位數的控制腳位 
#define pin_DIG4 15 // 第四位數的控制腳位
#define pin_A 13 // 七段顯示器 A 段 
#define pin_B 14 // 七段顯示器 B 段 
#define pin_C 26 // 七段顯示器 C 段 
#define pin_D 33 // 七段顯示器 D 段 
#define pin_E 25 // 七段顯示器 E 段 
#define pin_F 12 // 七段顯示器 F 段 
#define pin_G 27 // 七段顯示器 G 段 
#define pin_DP 32 // 七段顯示器小數點
// 共陽極設定 
#define hardwareConfig 1 // 定義硬體類型: 1 為共陽極, 0 為共陰極 
// 條件編譯: 根據硬體類型設定高低電平控制 
#if hardwareConfig==0 
#define digitOn 0 
#define digitOff 1
#define light 1 
#define dark 0 
#else 
#define digitOn 1 
#define digitOff 0 
#define light 0 
#define dark 1
#endif
// 儲存數字控制腳位 
byte digitPins[4] = {pin_DIG1, pin_DIG2, pin_DIG3, pin_DIG4}; // 四個位數腳位 

byte segmentPins[8] = {pin_A, pin_B, pin_C, pin_D, pin_E, pin_F, pin_G, pin_DP}; // 七段顯示器與小數點腳位
// 定義 0-9 的七段顯示器控制方式 
int num[10][8] = { {light, light, light, light, light, light, dark, dark}, // 0 
{dark, light, light, dark, dark, dark, dark, dark}, // 1 
{light, light, dark, light, light, dark, light, dark}, // 2
{light, light, light, light, dark, dark, light, dark}, // 3
{dark, light, light, dark, dark, light, light, dark}, // 4 
{light, dark, light, light, dark, light, light, dark}, // 5 
{light, dark, light, light, light, light, light, dark}, // 6 
{light, light, light, dark, dark, dark, dark, dark}, // 7
{light, light, light, light, light, light, light, dark},// 8
{light, light, light, light, dark, light, light, dark} // 9 };


// 顯示單個數字 
void displayNum(int pos, int number, bool showDot) 
{ 
for (int i = 0; i < 4; i++) 
{
digitalWrite(digitPins[i], digitOff); // 關閉所有位數 
} 
for (int i = 0; i < 8; i++) 
{ 
digitalWrite(segmentPins[i], num[number][i]); // 設定顯示的數字 
}
// 控制小數點顯示 
digitalWrite(segmentPins[7], showDot ? light : dark); digitalWrite(digitPins[pos], digitOn); // 啟動指定的位數 
}
// 顯示整個時間 (四個數字)
 int scanTime = 2; // 掃描間隔時間 (毫秒) 
void display(int num0, int num1, int num2, int num3) 
{ 
for (int i = 0; i < 50; i++)
{ 
// 快速掃描多次以持續顯示完整時間 
displayNum(0, num3, false); // 顯示第 1 位數 delay(scanTime); displayNum(1, num2, false); // 顯示第 2 位數 
delay(scanTime); displayNum(2, num1, true); // 顯示第 3 位數 (附小數點) 
delay(scanTime);
displayNum(3, num0, false); // 顯示第 4 位數 delay(scanTime); } }

// 初始化 
void setup() { 
Serial.begin(115200); 
// 連接 Wi-Fi
Serial.print("Connecting to "); 
Serial.println(ssid); 
WiFi.begin(ssid, password); 
while (WiFi.status() != WL_CONNECTED) 
{ 
delay(500);
Serial.print("."); 
} 
Serial.println("\nWiFi connected."); 
Serial.print("IP位址: "); 
Serial.println(WiFi.localIP());

Serial.print("WiFi RSSI: "); 
Serial.println(WiFi.RSSI()); 
// 初始化並同步 NTP 時間 
configTime(gmtOffset_sec, daylightOffset_sec, ntpServer); // 設定腳位為輸出模式 
for (int i = 0; i < 4; i++)
 { pinMode(digitPins[i], OUTPUT);
 digitalWrite(digitPins[i], digitOff); // 預設關閉 
} 
for (int i = 0; i < 8; i++)
 { pinMode(segmentPins[i], OUTPUT);
 digitalWrite(segmentPins[i], dark); // 預設為暗 
} 
}

// 主迴圈 
void loop() {
 struct tm timeinfo;
 if (getLocalTime(&timeinfo)) { // 獲取當前時間 
int num0 = timeinfo.tm_hour / 10; // 時的十位 
int num1 = timeinfo.tm_hour % 10; // 時的個位
int num2 = timeinfo.tm_min / 10; // 分的十位
int num3 = timeinfo.tm_min % 10; // 分的個位 display(num0, num1, num2, num3); // 顯示時間 
} 
else { Serial.println("Failed to obtain time"); // 獲取時間失敗提示
}
}
二、Python： 
from openpyxl import Workbook 
//含式載入 
wb = Workbook() 
//把檔案抓下來 ws = wb.active 
//活頁簿
 ws['A1'] = 42 
//座標

# Save the file 
wb.save("sample.xlsx")//儲存EXCEL檔案
