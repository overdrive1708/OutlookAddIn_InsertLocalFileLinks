# OutlookAddIn_InsertLocalFileLinks

## 【簡単な説明】

Outlookでローカルファイルリンクを簡単に挿入するためのアドインです。 

## 【開発の背景】

社内メールの場合、ファイルの場所をファイルサーバのパスでやり取りすることが多いのですが、以下の問題がありました。  

- ローカルファイルリンクにするための書式が周知されておらず、リンク切れが多発しており、メールからリンクをコピーしてエクスプローラーに貼り付けていました。
- リンクを挿入する機能はOutlook標準でもありますが、HTML形式にして、挿入→リンクで画面を開き、リンク先やパス、表示文字列などを入力する必要があり、手間でした。

そのため、フォルダ/フォルダをドラッグ&ドロップするだけで、簡単に所定の書式でローカルファイルリンクを挿入できるアドインを開発しました。

## 【必要要件】

- Microsoft .NET Framework 4.7.2
- Microsoft Visual Studio 2010 Tools for Office Runtime

不足している場合は、インストール時に自動的にインストールされます。

## 【インストール方法】

1. OutlookAddIn_InsertLocalFileLinks/installer/setup.exe」を実行します。
2. インストーラの指示に従ってインストールしてください。

## 【使用方法】

1. Outlookを起動してメールの編集画面を開きます。
1. リンクを挿入したい場所にカーソルを合わせます。
1. 「メッセージ」タブの、「ローカルファイルリンク挿入」という場所にある「挿入画面起動」をクリックします。
1. 挿入画面にリンクを挿入したいフォルダ/ファイルをドラッグ&ドロップします。

## 【設定による出力の違い】

「フォルダだけリンクにする」のチェックの有無によって、挿入結果が以下のように変わります。  
お好みで設定をしてください。  

■「フォルダだけリンクにする」をチェックしなかった場合  
W:\test  
    |----dirA(フォルダ)  
    |----fileA.txt  
の構成でdirAとfileA.txtをドラッグ&ドロップした場合、
```
<"file://W:\test\dirA">
<"file://W:\test\fileA.txt">
```
という内容が挿入されます。  

■「フォルダだけリンクにする」をチェックした場合  
W:\test  
    |----dirA(フォルダ)  
    |----fileA.txt  
の構成でdirAとfileA.txtをドラッグ&ドロップした場合、
```
<"file://W:\test">
・dirA
・fileA.txt
```
という内容が挿入されます。  
一つ上のフォルダだけリンクにして、中身はリンクにしない方法です。  

## 【アンインストール方法】

1. スタートメニューより、「設定」→「アプリ」→「アプリと機能」を開きます。
1. 「OutlookAddIn_InsertLocalFileLinks」をクリックして、「アンインストール」をクリックしてください。

## 【開発環境】
Microsoft Visual Studio Community 2019  
Version 16.8.2  

VisualStudio.16.Release/16.8.2+30717.126  
Microsoft .NET Framework  
Version 4.8.03752  

Office Developer Tools for Visual Studio   16.0.30502.00  
Microsoft Office Developer Tools for Visual Studio  

## 【ライセンス】

このプロジェクトはMITライセンスです。
詳細は [LICENSE](LICENSE) を参照してください。

## 【作者】

[overdrive1708](https://github.com/overdrive1708)

## 【変更履歴】

詳細は [CHANGELOG](CHANGELOG.md) を参照してください。