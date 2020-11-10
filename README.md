# AutoReportAppProject
日報をフォームで登録し、登録した日報データをもとに週報をテキストファイルで出力することができるアプリケーションである。


# DEMO
入力した日報データから週報テキストファイルを出力する。

以下、日報データ、出力する週報内容とする。

【日報データ】

1,2020/10/01,・test1,・test1,・test1

2,2020/10/02,・test2@NewLine・test2,・test2,・test2

3,2020/10/03,・test3,・test3@NewLine・test3,・test3

4,2020/10/04,・test4,・test4,・test4@NewLine・test4

5,2020/10/05,・test5@NewLine・test5,・test5@NewLine・test5,・test5

【出力する週報内容】

【日付】

2020/10/01

【実施内容】

・test1 

【翌日予定】

・test1 

【課題】

・test1


【日付】

2020/10/02

【実施内容】

・test2

・test2 

【翌日予定】

・test2 

【課題】

・test2


【日付】

2020/10/03

【実施内容】

・test3 

【翌日予定】

・test3 

【課題】

・test3

・test3


【日付】

2020/10/04

【実施内容】

・test4 

【翌日予定】

・test4

・test4 

【課題】

・test4


【日付】

2020/10/05

【実施内容】

・test5

・test5 

【翌日予定】

・test5 

【課題】

・test5

・test5


# Features
以下の特徴から日報データを加工しやすいため、登録した日報データを加工しメールなどアプリ以外の手段で活用するとよい。

・csvファイルに日報を登録したデータを保存している

・登録した日報データをもとに週報をテキストファイルに出力


# Requirement
* .NET Framework 4.7.2
* CsvHelper.15.0.5
* Microsoft.Bcl.AsyncInterfaces.1.1.0
* Microsoft.CSharp.4.5.0
* System.Runtime.CompilerServices.Unsafe.4.7.1
* System.Text.Encoding.CodePages.4.7.1
* System.Threading.Tasks.Extensions.4.5.2


# Installation
「Installer」フォルダ配下すべてのファイルをコピーし、「SetupAutoReportAppProject.msi」を実行してください。


# Usage
1.アプリケーションを起動。

2.入力日報フォームで日報を入力。

3.2で登録した日報データを日報データリストフォームで確認。

4-1.3で登録した日報データを日報データリストフォームで選択し、週報テキストファイルを出力。

4-2.3の日報データをcsvファイルで出力することも可能。

フォルダ構成
* Installer：インストーラーを格納。
* Source：アプリケーションソースを格納。
* Tool：リリースビルドしたファイルを格納。


# Note
Windowsアプリケーションになります。


# Author
* ShotaTsuji

機会があれば、上記アプリケーションの利用よろしくお願いします。

ご意見等あれば、Issuesでご連絡よろしくお願いします。
