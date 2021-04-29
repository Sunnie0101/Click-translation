# Click-translation
108學年第二學期
# 程式設計與實習(二)-期末專題報告
## 摘要：即時英文單字查詢、客製化單字測驗系統
本專題分為二部分，一為透過pyperclip擷取複製文字 ，並使用Google Translation API進行翻譯。二為透過翻譯紀錄，建構每日客製化的單字背誦表，以及測驗系統。
## 壹、使用說明：
### 程式一、擷取複製文字並翻譯(copied_text_translation.py)
1. 執行檔案
2. 將欲翻譯的文字反白並複製(ctrl+c)
![image](https://user-images.githubusercontent.com/60318542/116182068-42817480-a74e-11eb-824c-4105586384e3.png)
3. 彈跳視窗顯示翻譯內容
![image](https://user-images.githubusercontent.com/60318542/116182231-79f02100-a74e-11eb-96cc-2b679182873f.png)
4. (關閉視窗後)系統自動儲存紀錄
5. 重複2.~4.
6. 若不須使用則可直接結束程式

### 程式二、建構今日單字背誦表，以及測驗系統(words_exam_system.py)
1. 執行檔案
2. 如果尚未建立今日單字表，系統將自動生成
3. 開始測驗(分為模式一、模式二，詳細規則請見方法及流程圖)

|模式一|模式二|
|-----|------|
|![image](https://github.com/Sunnie0101/Click-translation/blob/main/img/words_exam_system_mode1.jpg)|![image](https://github.com/Sunnie0101/Click-translation/blob/main/img/words_exam_system_mode2.jpg)|

4. 若想暫時測驗，結束程式即可 (想繼續背誦則再次執行程式)
5. 當所有單字背誦完畢，則程式結束。

### 貳、方法及流程圖：
#### 一、Excel欄位介紹：
1. 單字總表(Vocabulary_list.xlsx)(粗正體字為實際儲存內容)

|	 |**voc**|**translation**|**value**|**last_test**|
|----|-------|-------|-----------------|---------------|
|	 |*英文單字*|*中文翻譯*|*記錄對單字的熟悉度*|*紀錄上次背誦情形*|
|*Ex:*|**sheet**|**片**|**11**|**Wrong**|
|	 |**desktop**|**桌面**|**8**|**0**|
|	 |**permanent**|**永久**|**10**|(None)|

(1) value計分規則(value越低，熟悉度越高)：

|原始value值|	~5|	6~10|	11~|
|-------|-------|-----------------|---------------|
|於模式一回答錯誤|	0|	+1|	0|
|於模式二回答錯誤|	0|	-1|	-1|
|完全正確	|-1	|-2	|-3|

(2)last_test:

|上次於模式一回答錯誤|	Wrong|
|上次於模式一回答正確|	0|
|未有背誦紀錄	|(None)|


2.今日背誦單字表(Vocabularyyyyy-mm-dd.xlsx)(粗正體字為實際儲存內容)

|voc|	translation	|mode|	situation|	value1|	value2|
|-------|-------|-----------------|---------------|---|---|
|英文單字|	中文翻譯|	背誦模式|	於模式二的背誦情形|	是否於模式一回答錯誤|	總背誦次數|
|sheet	|片|	0	|0	|False|	0|
|desktop|桌面|	1|	0|	(None)	|0|
|permanent	|永久|	2	|0	|False|	0|

(1) mode背誦模式:

|mode|	模式	|說明|
|--|--|--|
|0	|模式一|	題目：英文，為中文三選一選擇題|
|1	|模式二|	題目：中文，為英文填充題|
|2	|通過|

(2) situation於模式二的背誦情形：
初始值為0，答對加1，答錯減1

(3) value1是否於模式一回答錯誤：
初始值「(None)」，答錯為「False」

※(部分測驗規則參照’Quizlet’單字「學習」功能)  
    a. 當模式一正確後，轉為模式二  
    b. 當二個模式都正確，則通過  
    c. 若於模式二錯2題，則返回模式一，並且需再回到模式二時，需答對2題，才算通過

#### 二、程式簡介：
1. 擷取複製文字並翻譯(copied_text_translation.py)
使用函式庫：

|引入函式庫|	功能	|安裝方式|
|---|---|---|
|`import requests`|建立各種 HTTP 請求，從網頁伺服器上取得資料。|`pip3 install requests` |
|`import time`|時間函式庫	|(python內建)|
|`import pyperclip`|剪貼板(clipboard)功能|`pip install pyperclip`|
|`from tkinter import * `|GUI介面	|`pip install tkinter`|
|`import openpyxl`|Excel檔案讀寫|	`pip install openpyxl`(anaconda內建)|

程式流程圖：
![image](https://github.com/Sunnie0101/Click-translation/blob/main/img/copied_text_translation_flowchart.jpg)
2. 建構今日單字背誦表，以及測驗系統(words_exam_system.py)
使用函式庫：

|引入函式庫|	功能	|安裝方式|
|-------|-------|-------|
|`import openpyxl`|	Excel檔案讀寫|`pip insrall openpyxl`(anaconda內建)|
|`import datetime`|	日期時間函式庫	|(python內建)|
|`import random`|	剪貼板(clipboard)功能|`pip install random`(anaconda內建)|
|`import os`|	操作文件、目錄、系统…	|(python內建)|

程式流程圖：
![image](https://github.com/Sunnie0101/Click-translation/blob/main/img/words_exam_system_flowchart.jpg)
### 參、貢獻說明：
#### 一、將擷取clipboard資料與Google translation API：
參照文獻1.的getcopytext()以及文獻2的request API寫法，結合二功能。
#### 二、製作出客製化的每日背誦單字表，以及測驗系統

### 肆、未來展望：
#### 一、將翻譯方式由Google translation API轉成爬蟲的方式。由於API只提供翻譯，而沒有詞性、英英詞譯、例句…等，若是使用爬蟲的方式便能取得更多單字的資料。
#### 二、為測驗系統製作GUI介面
#### 三、將測驗系統的選字方式改為加權

### 伍、參考文獻：
1. 基于Python3.6写的自助翻译小软件--使用google translate的接口，Python实现爬取google翻译API结果，并打包成.exe的可执行文件 - WilsonSong1024 2018-05-13 15:16:29
https://blog.csdn.net/WilsonSong1024/article/details/80299335
2. 在python3中调用google 翻译api进行翻译(需要拥有google api密钥) - 猫哥的鱼库 2018-10-25 16:29:44
https://blog.csdn.net/qq_26870933/article/details/83381781?utm_medium=distribute.pc_relevant.none-task-blog-baidujs-2
3. Python笔记(十四）：操作excel openpyxl模块 - free赖权华 2018-07-04
https://cloud.tencent.com/developer/article/1156600
4. Python使用openpyxl读写excel文件的方法 – sunhaiyu 2017-06-30 09:51:10
https://www.jb51.net/article/117515.htm
5. 免費的學習工具和單詞卡Quizlet
https://quizlet.com/

