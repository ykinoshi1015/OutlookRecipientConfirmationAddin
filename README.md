# Outlook宛先表示アドイン


Outlook Recipient Confirmation Add-In、通称Outlook宛先表示アドイン。  

メールを送信する際の宛先表示機能を、デスクトップ版 Outlook2016で使えるよう開発されました。  


## 機能

Outlook宛先表示アドインを追加することにより、2つの機能が追加されます。  


### 機能1. メッセージ送信時の宛先確認

メッセージ送信時に以下のような宛先確認画面が表示されます。  

確認画面で宛先を確認することで、メッセージの誤送信を防ぎます。

![readme_feature1](https://user-images.githubusercontent.com/34431835/34712348-df3f9f34-f565-11e7-9b42-84501d3e45fa.PNG)

### 機能2. 受信トレイなどにあるメールの宛先確認

受信トレイをはじめ、下書きや送信済みアイテムフォルダにあるメールの宛先確認ができます。  

メールが、誰に宛てられたものなのかを「宛先リスト」のウィンドウより、確認できます。  

「宛先リスト」を開くには、宛先を確認したいメールを選択し、  
ホームタブ（または、メッセージ／会議タブなど）に表示される「宛先確認」のアイコンをクリックします。  

![github_ribbon](https://user-images.githubusercontent.com/29644865/38126302-b505fdc2-342a-11e8-9a57-774739608540.PNG)


（※受信したメールのBccは、表示されません）

## 必要条件

Outlook宛先表示アドインには、以下の環境が必要です。

*  Windows7またはWindows10
*  **デスクトップ版** Outlook2016
* アドレス帳にExchangeを利用していること

## インストール手順
### 確認事項

インストール実施前に、以下をご確認ください。  

*  インストールの実行をネットワークドライブ上で行わないこと
*  プロキシを使用していないこと  
* **Outlook2016を終了しておくこと**

また、以下のプログラムがお使いのPCにインストールされていない場合、自動的にインストールされます。

*  Microsoft .NET Framework 4.5.2
*  Windows インストーラー 4.5

### 手順
1. setupフォルダをクリックします  
![readme_installation_setupfolder](https://user-images.githubusercontent.com/34431835/34713370-6470b122-f569-11e7-91a5-a9a3107010cd.PNG)

2. setup.exeをクリックし、Downloadボタンを押します  
![readme_installation_setup zip](https://user-images.githubusercontent.com/29644865/34401117-aec2321a-ebdb-11e7-80c8-ef7945369371.PNG)

3. ダウンロードしたsetup.exeを実行します。

4. 詳細情報を押し、実行を選択します  
![readme_installation_protected](https://user-images.githubusercontent.com/29644865/36298153-f8a34402-1339-11e8-9694-b89422aab416.PNG)
![readme_installation_protected2](https://user-images.githubusercontent.com/29644865/36298152-f87d8fdc-1339-11e8-8480-c7abc3960d35.PNG)

5. この画面が表示されたら、Agreeボタンを押します  
![readme_streams_agreement](https://user-images.githubusercontent.com/29644865/36298004-222d0bc4-1339-11e8-8d60-8a0a80ee9e26.PNG)

6. インストールボタンを押します  
![readme_installation_check](https://user-images.githubusercontent.com/29644865/36297991-09ed738c-1339-11e8-9ed3-9ed8002ab965.PNG)

7. インストール完了です  
![readme_installation_completed](https://user-images.githubusercontent.com/29644865/36298698-ba952182-133c-11e8-8074-579497477808.PNG)


**お疲れ様でした**  :smiley:

Outlook2016を起動すると、アドインが有効になります。  


## FAQ

インストール時のエラーなど
[FAQ](https://github.com/ykinoshi1015/OutlookRecipientConfirmationAddin/wiki/FAQ)


## License

Copyright (c) 2018 Yuna Nakanishi / Yasuyuki Kinoshita / Kenta Kosaka   
Released under the MIT license   
[MIT-LICENSE.txt](./MIT-LICENSE.txt)   


