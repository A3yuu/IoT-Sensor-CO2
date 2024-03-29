# IoT Sensor
 
CO2,温度,湿度をM5 AtomとSCD4x、GASを使ってIoTするやつ。

# Features
![Product](https://user-images.githubusercontent.com/21051958/127768956-96f1382d-b6e6-4baf-ba8e-c4e9e34c0b28.jpg)

1万円以下で作れてとても小さい。
WEBブラウザからどこでもグラフ表示で履歴確認できる。
GAS連携のベースサンプルとしてどうぞ。

# Requirement Hardware

* Wireless Router

* M5 Atom

https://www.switch-science.com/catalog/6262/

* SCD4x

https://www.switch-science.com/catalog/7168/

* ATOMICプロトキット

https://www.switch-science.com/catalog/6345/

* 電子工作道具一式

PC/USB-Cケーブル/半田ごて/ビニル線/ピンセット/ワイヤストリッパ

SCD4xピッチ変換使わないなら：耐熱テープ/スズメッキ線

# Requirement Software

* Arduino IDE
https://www.arduino.cc/en/software

* M5Atom Library

* FastLED Library

# Installation Hardware

次のように、対応するピンを接続する。

ATOMICプロトキット:SCD4x

3V3:VDD,VDDH

GND:GND×6

SCL:SCL

SDA:SDA

![Back](https://user-images.githubusercontent.com/21051958/127768958-fe3138f1-06a4-40b3-8a61-20056b50221d.jpg)

![Front](https://user-images.githubusercontent.com/21051958/127768959-582dd3c7-a23a-4c17-89ff-922ec55ee83c.jpg)

# Installation Software

�@Arduino IDEで「IoTSensor/IoTSensor.ino」を書き込む

�AGoogle Spreadsheetを作成する。

�BSpreadsheetを公開する。

�CSpreadsheetのツール内「スクリプト エディタ」に「scripts/code.gs」「scripts/index.html」を書き、デプロイする。

# Usage

プロダクトにUSB電源を入れる。

ボタン押しでWPSモードになるので、インターネットに繋がったルーターにWPSで接続する。

同じサブネット内の端末からブラウザアクセスする。

デプロイしたGASのアドレス「 https://script.google.com/macros/s/***/exec 」を入力し、SETボタンを押す。

<img width="376" alt="localpage" src="https://user-images.githubusercontent.com/21051958/127769613-1b0a9ae7-d988-479b-a48f-b0022bb01c84.png">

デプロイしたGASのアドレス「 https://script.google.com/macros/s/***/exec 」にブラウザからアクセスする。
 
<img width="513" alt="webpage" src="https://user-images.githubusercontent.com/21051958/127769615-5d74866d-4155-4d58-98a6-43576397ad7c.png">

# Usage LP

Low Power版はSCD41でのみ有効で、省電力で熱の発生が少ないので温湿度精度が少しいいやつ。

プロダクトにUSB電源を入れる。

ボタン押しで起動(リセット)でLEDが青で起動。10分間サーバーモードとなり、ブラウザアクセスができるようになる。

サーバーモードでボタンを押すとWPSモードになるので、インターネットに繋がったルーターにWPSで接続する。

以下通常モードと同じ。

# Usage 2

第2版はSGP40を追加したもので、VOC指標が追加取得できるやつ。

スクリプトはindex_vocの方をindexとして適用する必要がある。

# Note

下記参考

SCD4xデータシート
SGP40データシート

https://github.com/Sensirion/embedded-sgp

URLのデコード(https://qiita.com/dojyorin/items/07efd4a55039bcc30cb2)

# Author

https://twitter.com/A3_yuu

https://github.com/A3yuu/IoT-Sensor

# License

(c)2021 A3_yuu

"IoT Sensor" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).
