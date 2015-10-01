# PowerPoint Auto Update Total Page
<h1>POWERPOINT 自動更新總頁數教學</h1>
<h2>步驟:</h2>
<ol>
<li>開啟PPT->檔案->選項->自訂功能區->開發人員。</li>
<li>檢視->投影片母片->開發人員->LABEL->右鍵->檢視程式碼。</li>
<li>貼上下面程式碼。</li>
<li>完成。</li>
</ol>
</br>
```
Sub OnSlideShowPageChange()
    For i = 1 To ActivePresentation.Slides.Count
        Label1.Caption = "  " & i
        Label1.Font = "Times New Roman"
        Label1.Font.Size = 12
    Next
End Sub
```
<h2>效果:</h2>

每次修完PPT，按播放後總頁數會自動更新。
