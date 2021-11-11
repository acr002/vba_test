## ファイル名の取得
VBAでファイル名を取得する方法はいくつかありますが、私がもっともよく使うものを備忘録として記述します。

下記の関数にフォルダを指定して呼び出します。
フルパスが入ったCollectionを返します。

```VB
Private Function filenames_sub(ByVal a_path As String) As Collection
  Dim fso      As Object
  Dim r_cc     As Collection
  Dim cc       As Collection
  Dim ii       As Variant
  Dim b_file   As Object
  Dim b_folder As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set cc = New Collection
  For Each b_file In fso.getfolder(a_path).Files
    cc.Add b_file.path
  Next b_file
  For Each b_folder In fso.getfolder(a_path).subfolders
    Set r_cc = filenames_sub(b_folder.path)
    For Each ii In r_cc
      cc.Add ii
    Next ii
  Next b_folder
  Set fso = Nothing
  Set filenames_sub = cc
End Function
```

## この関数について
プライベートな関数としています。
最初はクラスにしていたのですが、移植等が面倒なので関数にしました。

使用する機能はFileSystemObjectです。
呼び出しは参照設定ではなくCreateobjectにしています。
速度より配布を考慮しました。

再帰的に処理します。
ファイル名(フルパス)をCollectionに格納し、そのCollectionを返します。
Collectionを使用したことのない方は身構えてしまうかもしれませんが、中身は文字列です。
関数からの戻り値を受け取る際はSetステートメントを使用します(おまじないと思ってください)。

## 呼び出し元
関数は呼び出して使用しますので、呼び出し元が必要です。
また、この関数の引数は対象のフォルダです。「どのフォルダを調べるか」を指定するということです。
フォルダの指定方法について、２つサンプルを作りました。

### 対象フォルダをコード内で指定する方法
このマクロ入りのエクセルが置いてあるフォルダに「in」というフォルダがあると仮定します。
(「in」フォルダがないとエラーになります。)
エクセルが置いてあるところは`ThisWorkbook.path`で取得します。
その取得したパスに`\in\`を追加して完成です。
(inのところを自由に変えて使用してください)

私はこの方法を使用しています。
とりあえず「in」というフォルダを作って、その中にフォルダごとでいいので対象のファイルを入れてしまえばいいからです。
「in」フォルダというルールが設けられるのであれば使い勝手はいいと思います。

```VB
Public Sub sample1_on_code()
  Dim fns  As Collection
  Dim path As String
  path = ThisWorkbook.path & "\in\"     '指定するフォルダ
  Set fns = filenames_sub(path)     '関数を呼び出してfnsに結果を入れます
End Sub
'-----------------------------------------------------------------------------
```

### ダイアログを使う
マクロを実行する人が誰かわからない場合や、どんな状況で使用されるかわからない場合のためのダイアログを使う方法です。
マクロを実行すると、エクセルがあるフォルダを基準にしてダイアログが開きます。
フォルダを指定してもらえば、関数に指定されたフォルダを渡します。
フォルダが指定されなかったら基準のフォルダを関数に渡します。

```VB
Public Sub sample2_dialog()
  Dim path_this As String
  Dim path      As String
  Dim fns       As Collection
  path_this = ThisWorkbook.path & "\"
  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = path_this
    .InitialView = msoFileDialogViewDetails
    .Title = "『フォルダ』を選んでください"
    If .Show = True Then
      path = .SelectedItems(1)
    Else
      path = path_this
    End If
  End With
  Set fns = filenames_sub(path)
End Sub
```


## その他の方法
ブックを開く際にはファイル名が必要となります。
一つだけを対象にするならVBAの記述内に直接書く事も一つの手です。
取得ではなく、指定ですが時間がなかったり、自分しか使わない場合はこの方法でもいいと思います。

欲しいファイル名がいくつもある場合は、ワイルドカードが使えるDir関数も便利です。
特に一つのフォルダ内だけで、階層もないような条件の場合は、気軽さから候補になると思います。
ただ、Dir関数には注意点があります。それは次のファイル名を取得するには「引数を省略して呼び出す」ということです。
メインルーチンでDir関数を使用し、いくつか飛んだ先のルーチンでもDir関数を使いたいと思い引数を指定してしまうと、メインルーチンのDir関数も変化してしまいます。
うまく回避する方法があるかもしれませんが、私は知らないので、Dir関数とは距離を取っています。
また、Do loopを使わなければならないこともDir関数を使わない理由の一つです。

