# Army_LINE_AutoReport
國軍休假自動回報排程  


一個使用selenium自動化套件操作LINE Chrome Extension的自動回報程式, 讓你完完整整的享受休假時間  

## 檔案說明  
### source_code (python源代碼)
* ROC_armed_forces_auto_report.ipynb
* ROC_armed_forces_auto_report.py
* chromedriver.exe         
* extension_2_4_1_0.crx

### executable (已使用pyinstaller打包成exe檔, 可直接執行)
* ROC_armed_forces_auto_report.exe
    * 主程式
* chromedriver.exe         
    * selenium使用的webdriver (目前使用ChromeDriver 88.0.4324.96)
    *  webdriver版本需與Chrome版本相同  可至 https://chromedriver.chromium.org/downloads 下載新版本
* extension_2_4_1_0.crx
    * LINE擴充元件
  
## 使用說明
* 下載解壓縮後，點擊executable\ROC_armed_forces_auto_report.exe
* 至設定->系統->電源與睡眠 將睡眠時間更改為永不
* 本程式所使用chromedriver.exe需與當前電腦Chrome版本相同，可至Chrome瀏覽器右上角標示 :->說明->關於Google Chrome 找到當前使用的版本
  
## 介面說明 
![example](https://user-images.githubusercontent.com/48814609/110164957-16700400-7e2d-11eb-9edd-e8589a7ff92b.PNG)
