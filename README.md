# vba
Sub sample()

　Dim objIE As InternetExplorer

　'IE(InternetExplorer)のオブジェクトを作成する
　Set objIE = CreateObject("InternetExplorer.Application")

  'IE(InternetExplorer)を表示する
  objIE.Visible = True

  '指定したURLのページを表示する
  objIE.Navigate "http://www.vba-ie.net/"

　'完全にページが表示されるまで待機する
　Do While objIE.Busy = True Or objIE.ReadyState <> 4
　　DoEvents
　Loop

End Sub
