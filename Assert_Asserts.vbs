Option Explicit

Private Sub Include(ByVal FileName)
	Const ForReading = 1
	Dim fso: set fso = CreateObject("Scripting.FileSystemObject")   
	Dim f: set f = fso.OpenTextFile(FileName,ForReading)
	Dim s: s = f.ReadAll()
	ExecuteGlobal s
	
	Set fso = Nothing
	Set f = Nothing
	Set s = Nothing
End Sub

Include("Assert.vbs")

Dim Assert: Set Assert = new Assert_
Dim Be: Set Be = new Be_
Dim Text: Set Text = new Text_

Dim BeToo: Set BeToo = Be
Dim NotIntialized: Set NotIntialized = Nothing

Assert.That BeToo, Be.SameAs(Be)
Assert.That 1, Be.GreaterThan(0)
Assert.That 0, Be.LessThan(1)
Assert.That 1, Be.GreaterThanOrEqualTo(0)
Assert.That 1, Be.GreaterThanOrEqualTo(1)
Assert.That 1, Be.AtLeast(0)
Assert.That 1, Be.AtLeast(1)
Assert.That 0, Be.LessThan(1)
Assert.That 0, Be.LessThanOrEqualTo(0)
Assert.That 0, Be.AtMost(1)
Assert.That 0, Be.AtMost(0)
Assert.That Assert, Be.TypeOf_(new Assert_)
Assert.That NotIntialized, Be.Null_()
Assert.That "test", Be.NaN()
Assert.That "1234", Text.StartsWith("123")
Assert.That "1234", Text.EndsWith("34")
Assert.That "1234", Text.Matches("\d+")
Assert.That "1234", Text.Contains("23")

Set Assert = Nothing
Set Be = Nothing
Set Text = Nothing

wscript.Echo "All assertions were valid!"
