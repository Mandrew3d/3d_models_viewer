' update_models.vbs
' Автоматически добавляет новые .glb модели в viewer.html
' Работает при двойном клике, без консоли

Option Explicit

Dim fso, htmlPath, modelPath, htmlText, modelFile, modelName
Dim htmlFile, modelFolder, linkPattern, matches, existingLinks
Dim newLinks, linkText, updatedHtml, openTag, closeTag

Set fso = CreateObject("Scripting.FileSystemObject")

htmlPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "viewer.html")
modelPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "model")

If Not fso.FileExists(htmlPath) Then
  MsgBox "❌ Не найден viewer.html рядом со скриптом.", vbCritical, "Ошибка"
  WScript.Quit
End If

If Not fso.FolderExists(modelPath) Then
  MsgBox "❌ Не найдена папка model рядом со скриптом.", vbCritical, "Ошибка"
  WScript.Quit
End If

' Читаем HTML
Set htmlFile = fso.OpenTextFile(htmlPath, 1, False, 0)
htmlText = htmlFile.ReadAll
htmlFile.Close

' Ищем существующие ссылки
Set existingLinks = CreateObject("Scripting.Dictionary")
Dim regEx, matchesObj, m
Set regEx = New RegExp
regEx.Pattern = "href=""viewer/viewer\.html\?m=([^""]+)"""
regEx.Global = True

Set matchesObj = regEx.Execute(htmlText)
For Each m In matchesObj
  existingLinks.Add m.SubMatches(0), True
Next

' Ищем .glb файлы
Set modelFolder = fso.GetFolder(modelPath)
newLinks = ""

For Each modelFile In modelFolder.Files
  modelName = fso.GetFileName(modelFile)
  If LCase(Right(modelName, 4)) = ".glb" Then
    If Not existingLinks.Exists(modelName) Then
      linkText = "    <a href=""viewer/viewer.html?m=" & modelName & """ class=""model-link"">🧩 " & Replace(modelName, ".glb", "") & "</a>" & vbCrLf
      newLinks = newLinks & linkText
    End If
  End If
Next

If newLinks = "" Then
  MsgBox "✅ Новых моделей нет — index.html уже актуален.", vbInformation, "Готово"
  WScript.Quit
End If

' Вставляем новые ссылки в HTML
Dim patternOpen, patternClose, startPos, endPos
patternOpen = InStr(htmlText, "<div id=""modelList""")
patternClose = InStr(htmlText, "</div>")
If patternOpen = 0 Or patternClose = 0 Then
  MsgBox "❌ Не удалось найти блок <div id=""modelList""> в index.html", vbCritical, "Ошибка"
  WScript.Quit
End If

Dim insertPos
insertPos = patternClose - 1
updatedHtml = Left(htmlText, insertPos - 1) & vbCrLf & newLinks & Mid(htmlText, insertPos)

' Сохраняем
Set htmlFile = fso.OpenTextFile(htmlPath, 2, False, 0)
htmlFile.Write updatedHtml
htmlFile.Close

MsgBox "✅ Добавлено новых моделей: " & CountLines(newLinks), vbInformation, "Успех"

' Подсчёт строк
Function CountLines(s)
  Dim lines
  lines = Split(Trim(s), vbCrLf)
  CountLines = UBound(lines) + 1
End Function
