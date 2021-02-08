# uquery
VBAで ADO をもっと簡単に使う

### コードの例
```vba
Option Explicit

Public con As New UConnection

Sub ボタン_Click()
    Dim q As New UQuery

    q.SetCon con
    q.Sql.Add "SELECT * FROM [ごはん$]"
    q.Sql.Add "WHERE 価格 >= ?"
    q.AddParam "価格", 300
    q.OpenRecordset

    Do Until q.hasNext
        MsgBox q.Fields("ごはん") & " ￥" & q.Fields("価格")
        q.MoveNext
    Loop
End Sub
```
### サンプルイメージ
![sample1](https://cdn-ak.f.st-hatena.com/images/fotolife/u/uhoo/20181123/20181123162004.gif)
