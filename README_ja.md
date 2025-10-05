# VBAForm2Tkinter - Excel VBA UserForm to Tkinter Converter
🌎[English](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/blob/main/README.md)<br><br>
このプログラムは、Excel VBAにて作成したユーザーフォームをPythonのTkinter用に変換可能なプログラムです<br>

## 変換例
<img width="681" height="1275" alt="Image" src="https://github.com/user-attachments/assets/9bd8ae88-4f31-411d-8acf-e83fc5c235b6" /><br>
<img width="704" height="695" alt="Image" src="https://github.com/user-attachments/assets/45678575-a162-4e18-a516-ba522727d1f9" /><br><br>

## 動作要件
- 対応OS: Windows
- 必要ソフトウェア: Microsoft Excel

## 動作確認済環境
= Windows 10/11
- Excel 2010(32bit)
- Excel 2016(32bit)
- Excel 2019(64bit)

## 反映する項目
- 変数名(オブジェクト名)
- コントロールのおおよそのレイアウトとサイズ
- コントロールの色(文字色、背景色)
- テキスト表示(Label, CommandButton, CheckBox, ToggleButton, OptionButton, MultiPage)
- フォント(フォント種類、サイズ、太字、斜体)
- 枠線(UserForm, Frame, TextBox, Label, ListBox, Image)
- マウスカーソル
- テキスト表示の左寄せ・中央・右寄せ(Label, TextBox [MultiLine=False], ComboBox, CheckBox, ToggleButton, OptionButton)
- TextBox, ComboBoxのデフォルト値
- ComboBox, ListBoxに設定したアイテム
- OptionButton, CheckBox, ToggleButtonの選択状態
- BackStyleに設定した透明表示設定

## 対応しているコントロールの種類
| VBA Formのクラス | Tkinterのクラス|
| ------ | ------ |
| Label | tk.Label |
| CommandButton | tk.Button |
| Frame (Captionなし) | tk.Frame |
| Frame (Captionあり) | tk.LabelFrame |
| TextBox (MultiLine=False) | tk.Entry |
| TextBox (MultiLine=True) | tk.Text |
| SpinButton | tk.Spinbox |
| ListBox | tk.Listbox |
| CheckBox | tk.Checkbutton |
| ToggleButton | tk.Checkbutton(indicatoron=0) |
| OptionButton | tk.Radiobutton |
| Image | tk.Canvas |
| ScrollBar | ttk.Scale |
| ComboBox | ttk.Combobox |
| MultiPage | ttk.Notebook |

※SpinButtonは仕様が異なるため、配置方法によっては外観が異なります<br>
※ScrollBarについてはVBAのフォームには上下調整用のボタンがありますがTkinterのScaleにはありません<br>
<br>
上記以外のコントロールがフォーム上にある場合、変換に失敗するので該当のコントロールを削除したうえで再度変換を行ってください<br>

## 使い方
使用前に、変換したいユーザーフォームが作成されたExcelブックを用意する必要があります<br>
また、VBE上でイミディエイトウィンドウが表示されていない場合は表示の設定を行ってください<br><br>
<img width="843" height="768" alt="Image" src="https://github.com/user-attachments/assets/676cd54c-d610-4c25-bd9a-9e064e38dc5e" /><br><br>
1.[ここ](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/releases)から最新版のファイルをダウンロードし解凍してください、中のVBAForm2Tkinter.basを使用します<br>
2. Excelの開発→Visual BasicからVBEを開いてください<br>
3. プロジェクトを右クリックし、「ファイルのインポート」よりVBAForm2Tkinter.basをインポートします<br>
4. イミディエイトウィンドウに「Call ConvertForm2Tkinter(UserForm1)」と入力しEnterキーを押下します<br>
```vb
Call ConvertForm2Tkinter(UserForm1)
```
※「UserForm1」の部分は変換したいユーザーフォームのオブジェクト名に変えてください<br>
5. 正常に変換が完了した場合、メッセージが表示されExcelブックと同じディレクトリに「output.py」が作成されます<br>
6. GUIの外観を確認したら、pyファイルを編集し.mainloop()の上にコントロール名.configure()でボタン押下時の関数の設定などをしてください<br>

## 子要素を設定できないコントロールの並び順について
Tkinterでは例としてLabelにLabelを重ねた場合は設置した順番が後のものが優先して前面に表示されます<br>
ただしVBAのユーザーフォームにおいては前面/背面を変更することができるためこの限りではありません<br>
このプログラムは各コントロールを階層順にソート後、同じ階層のものについては元々の設置順に従いウィジェットを配置します<br>
現状コントロールのZオーダー(前面/背面情報)を取得できる手段がないため反映させることができずVBAでの表示と異なってしまう場合があります<br>
その場合は、Pythonのコードを編集し、前面に表示したいものを後に設置するか、VBA側でコントロールを配置し直すことで順番を後にしてください<br>
なお、新規でGUIを作成する場合は重ねるよりもFrameなどの明確な親子関係を設定可能なコントロールを使用することを推奨します<br>

## 使用のさいの注意点
マルチモニター環境でこのプログラムを使用する場合、一時的にモニターを1つにするか、すべてのモニターの拡大率を統一したうえで使用してください<br>
異なる拡大率のモニターが混在している場合、ウィンドウサイズの計算が正常に行えない可能性があります<br>

## 日本語コメントのソースコードについて
ソースコードのコメントは全文英語で記載していますが[ここ](https://gist.github.com/ZeeZeX/1f0bb62d9e476b0df2aed8653ca303d4)から日本語版を確認可能です、フォークするさいに必要であれば役に立ててください
