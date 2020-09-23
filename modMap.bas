Attribute VB_Name = "modMap"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long

Public Const TileWidth = 18
Public Const TileHeight = 18

Enum nDir
up = 0
down = 2
Left = 3
Right = 1
End Enum

Enum nAI_kinds
Still = 0
Pace = 1
Spin = 2
Random = 3
End Enum

Type PackItem
Caption As String
Value As Long
FuncID As String 'the function is preforms (ie magic dust = +10 magic, or candy = +5 health)
Act As Long
End Type

Type Pack
Item(1 To 20) As PackItem

mDust As Long
money As Long
End Type

Type BattleInfo
Health As Long
Attack As Long
Defense As Long
MaxHealth As Long
Level As Long

Exp As Long
TExp As Long
End Type

Type Player
X As Long
Y As Long

SpeachI As Long
Walking As Boolean
DestX As Long
DestY As Long
WalkC As Long

dir As nDir

B As BattleInfo
Pack As Pack

Step As Long
StepC As Long

SpellCut As Boolean
SpellFreeze As Boolean
End Type

Type Tile
X As Long
Y As Long

Type As Long

Walkable As Long
DoorPath As String
DoorWay As Long
DoorDir As Long

SignText As String
isBattle As Long
isMarket As Long
plID As Long

Ani As Long
AniC As Long
End Type

Type Map
Name As String
Midi As String
Mapwidth As Long
Mapheight As Long

MarketPath As String

MaxHP As Long 'the maxium stats of a wild monster in the high grass
MaxAttack As Long
MaxDefense As Long
End Type

Type obj
AI As Long
T As Long
Name As String

Speek() As String
SpeekI As Long

AI_Type As nAI_kinds
dir As Long
Person As Long

Walking As Long
WalkC As Long

X As Long
Y As Long
DestX As Long
DestY As Long
Step As Long
StepC As Long

plID As Long

Act As Long
End Type

Type enemy
HP As Long
MaxHP As Long

Attack As Long
Defense As Long

Name As String
End Type

Type MessageObject
message As String
Time As Long
End Type

Public M As Map
Public T(5000) As Tile
Public p As Player
Public o(1 To 20) As obj
Public Msg As MessageObject
Public Enem As enemy
Public E As New ClsEngine
Public FPS As Long
Public Louis As Long 'to determine whether you beat louis
Public gSpeed As Long
Public Night As Long 'determine whether it is day or night
Public Boss(1 To 5) As Long
Public TFPS As Long

Function LoadMap(pth As String, dr As Long, d)
Dim lineread() As String, C() As String, free, tfree, X, Y, func() As String, I
Dim funct As String, u() As String, path, door, dir
door = dr
dir = d
path = pth

ClearStuff

Msg.Time = 0
Msg.message = ""

Open App.path + "\maps\" & path For Input As #1

Input #1, M.Name
Input #1, M.MarketPath
Input #1, free, tfree

frmGame.market.Picture = LoadPicture(App.path + "\other data\" & free)
frmGame.battle.Picture = LoadPicture(App.path + "\other data\" & tfree)

Input #1, M.MaxHP, M.MaxAttack, M.MaxDefense
Input #1, M.Midi
Input #1, M.Mapwidth, M.Mapheight

tfree = ""
free = ""

Do Until EOF(1)

    Line Input #1, free

    If free = "END TILES" Then GoTo DoneTiles

    tfree = tfree & free & vbCrLf

Loop

DoneTiles:

I = 1

Do

    Line Input #1, funct

    If funct = "END OBJECTS" Then GoTo DoneObjs

        u() = Split(funct, ",")
        o(I).Act = True
        o(I).dir = nDir.down
        o(I).Name = u(0)
        o(I).Speek() = Split(u(1), "|")
        o(I).SpeekI = 0
        o(I).dir = Val(u(5))
        o(I).T = Val(u(2))
        o(I).AI = 0
        o(I).AI_Type = Val(u(4))
        o(I).plID = Val(u(3))
        o(I).Person = False
        If isValue(o(I).T, 1, 2, 3) = True Then o(I).Person = True

        If o(I).Name = "Louis" And Louis = True Then
        o(I).Speek() = Split("We'll youve beaten me|Bet if we had a rematch I'd win!", "|")
        o(I).Walking = False
        End If
        
    DoEvents
    I = I + 1

Loop

DoneObjs:

lineread() = Split(tfree, vbCrLf)
I = 0

frmGame.buffer.Width = M.Mapwidth * TileWidth
frmGame.buffer.Height = M.Mapheight * TileHeight

For Y = 1 To M.Mapheight
    
    C() = Split(lineread(Y - 1), "|")
    
    For X = 1 To M.Mapwidth
    
        func() = Split(C(X - 1), ",")
        
        T(I).X = (X - 1) * TileWidth
        T(I).Y = (Y - 1) * TileHeight
        T(I).Type = Val(func(0))
        T(I).DoorPath = func(1)
        If InStr(1, func(1), ":") > 1 Then T(I).DoorPath = Split(func(1), ":")(0)
        If Len(T(I).DoorPath) > 1 Then T(I).DoorWay = Val(Split(func(1), ":")(1))
        If Len(T(I).DoorPath) > 1 Then T(I).DoorDir = Val(Split(func(1), ":")(2))
        T(I).SignText = func(2)
        T(I).isMarket = Val(func(3))
        T(I).plID = Val(func(4))
        If T(I).plID = door Then p.X = T(I).X: p.Y = T(I).Y: p.DestX = p.X: p.DestY = p.Y
        T(I).Walkable = True
        If isValue(T(I).Type, 0, 2, 3, 7, 9, 10, 11, 12, 13) = True Then T(I).Walkable = False
        
        I = I + 1
    Next X
    
Next Y

Close #1

SetObjs

p.dir = dir
End Function

Function WalkOn(TileI As Long)
If T(TileI).DoorPath <> "" Then LoadMap T(TileI).DoorPath, T(TileI).DoorWay, T(TileI).DoorDir
If T(TileI).Type = 6 And Int(Rnd * 10) = 5 Then LoadBattle
End Function

Function GetTile() As Long
Dim I As Long

For I = 0 To M.Mapheight * M.Mapwidth
    
    If T(I).X = p.DestX And T(I).Y = p.DestY Then GetTile = I: Exit For

Next I
End Function

Function GetObjTile(Ind) As Long
Dim I As Long

For I = 0 To M.Mapheight * M.Mapwidth

    If T(I).X = o(Ind).DestX And T(I).Y = o(Ind).DestY Then GetObjTile = I: Exit For

Next I
End Function

Function SetObjs() As Long
Dim I As Long, i2

For I = 0 To M.Mapheight * M.Mapwidth

     For i2 = 1 To 20
     
        If o(i2).plID = T(I).plID Then
            o(i2).Walking = False
            o(i2).X = T(I).X
            o(i2).Y = T(I).Y
            o(i2).DestX = o(i2).X
            o(i2).DestY = o(i2).Y
            
            If o(i2).Name = "Louis" And Louis = True Then
                o(i2).X = o(i2).X - 18
                o(i2).DestX = o(i2).X
            End If
        
        End If
     
     Next i2
     
Next I
End Function

Function TalkTo()
If p.SpeachI > 0 Then Exit Function
p.SpeachI = 100

Dim TileI As Long

Select Case p.dir

    Case nDir.up: TileI = GetTile - M.Mapheight
    Case nDir.down: TileI = GetTile + M.Mapheight
    Case nDir.Left: TileI = GetTile - 1
    Case nDir.Right: TileI = GetTile + 1

End Select

If T(TileI).SignText <> "" Then DoMsg T(TileI).SignText

If T(TileI).isMarket = -1 Then LoadMarket

If isTileTaken(TileI) = True And Msg.Time < 11 Then
    
    Dim H As Long, odir

    H = GetObj(TileI)

    If o(H).Person = True Then
        
        Select Case p.dir
        
            Case nDir.down: odir = nDir.up
            Case nDir.Left: odir = nDir.Right
            Case nDir.Right: odir = nDir.Left
            Case nDir.up: odir = nDir.down
            
        End Select

    o(H).dir = odir
    
End If

    DoMsg o(H).Name & ": " & o(H).Speek(o(H).SpeekI)

    If Msg.message = "Old Man: Here I'll teach you the cut spell" Then p.SpellCut = True
    If Msg.message = "Louis: Let's Fight!" Then DoMsg "Fight me you chicken!": LoadSetBattle "Louis", 50, 18, 10, frmGame.bsLouis
    If Msg.message = "I will teach the freeze spell to you" Then p.SpellFreeze = True
    
    o(H).SpeekI = o(H).SpeekI + 1: If o(H).SpeekI > UBound(o(H).Speek()) Then o(H).SpeekI = 0

End If
End Function

Function ClearStuff()
Dim I As Long

For I = 0 To 5000

    T(I).X = 0
    T(I).Y = 0
    T(I).Type = 0
    T(I).plID = 0
    T(I).Ani = 0
    T(I).isMarket = False
    T(I).DoorDir = 0
    T(I).DoorPath = ""
    T(I).DoorWay = 0
    T(I).SignText = ""
    
Next I

For I = 1 To 20

    o(I).Act = False
    o(I).plID = 0
    
Next I
End Function

Function isValue(Value As Long, ParamArray Nums() As Variant) As Boolean
Dim I As Long

isValue = False

For I = LBound(Nums()) To UBound(Nums())

    If Value = Nums(I) Then isValue = True
    
Next I
End Function

Function CheckInput()
If E.IsPressed(vbKeyLeft) Then

    If p.Walking = False Then p.DestX = p.X - TileWidth: p.Walking = True
    
End If

If E.IsPressed(vbKeyRight) Then

    If p.Walking = False Then p.DestX = p.X + TileWidth: p.Walking = True
    
End If

If E.IsPressed(vbKeyUp) Then

    If p.Walking = False Then p.DestY = p.Y - TileHeight: p.Walking = True
    
End If

If E.IsPressed(vbKeyDown) Then

    If p.Walking = False Then p.DestY = p.Y - TileHeight: p.Walking = True
    
End If
End Function

Function LoadMenu()
Dim I As Long

frmGame.board.Visible = False
frmGame.menu.Visible = True
frmGame.battle.Visible = False
frmGame.main.Visible = False
frmGame.market.Visible = False

frmGame.lstItems.Clear

For I = 1 To 20

    If p.Pack.Item(I).Act = True Then
    
        frmGame.lstItems.AddItem p.Pack.Item(I).Caption
        
    End If

Next I
End Function

Function UseItem(cap As String, I)
Select Case UCase(cap)
    Case "CANDY" 'candy
        p.B.Health = p.B.Health + 5
    Case "MAGIC DUST" 'magic dust
        p.Pack.mDust = p.Pack.mDust + 10
    Case "ENERGY JUICE" 'energy juice
        p.B.Attack = p.B.Attack + 5
    Case "IRON JUICE" 'iron juice
        p.B.Defense = p.B.Defense + 5
    Case "SPECIAL CANDY"
        p.B.Level = p.B.Level + 1
        p.B.Attack = p.B.Attack + (Int(Rnd * 7) + 1)
        p.B.Defense = p.B.Defense + (Int(Rnd * 5) + 1)
End Select

p.Pack.Item(I + 1).Act = False
End Function

Function LoadBattle()
frmGame.tmrAttack.Enabled = False
frmGame.tmrAttack2 = False
frmGame.tmrMagic.Enabled = False
frmGame.imgMagic.Visible = False
frmGame.imgAttack.Visible = False
frmGame.imgAttack2.Visible = False
frmGame.imgEnemy.Picture = frmGame.enemy(Int(Rnd * (frmGame.enemy.ubound + 1))).Picture
frmGame.board.Visible = False
frmGame.menu.Visible = False
frmGame.battle.Visible = True
frmGame.main.Visible = False
frmGame.market.Visible = False

Enem.HP = Int(Rnd * M.MaxHP) + 1
Enem.MaxHP = Enem.HP
Enem.Attack = Int(Rnd * M.MaxAttack) + 1
Enem.Defense = Int(Rnd * M.MaxDefense) + 1
End Function

Function NewGame()
LoadMap "house.dat", 2, 2 'load the house map

p.dir = down

p.Pack.Item(1).Act = True
p.Pack.Item(1).Value = 50
p.Pack.Item(1).Caption = "Candy"
p.Pack.Item(2).Act = True
p.Pack.Item(2).Value = 50
p.Pack.Item(2).Caption = "Candy"
p.Pack.Item(3).Act = True
p.Pack.Item(3).Value = 100
p.Pack.Item(3).Caption = "Magic Dust"

p.SpellCut = False
p.SpellFreeze = False

p.B.Level = 1
p.B.Attack = Val(frmGame.txtAttack.text)
p.B.Defense = Val(frmGame.txtDefense.text)
p.B.Health = 50
p.B.MaxHealth = 50
p.B.Exp = 0
p.B.TExp = 10

p.Pack.money = 300
p.Pack.mDust = 30
p.Walking = False
p.DestX = p.X
p.DestY = p.Y
End Function

Function MovePlayer()
If p.Walking = True And p.WalkC > 3 Then
    p.WalkC = 0

    Select Case p.X
        Case Is < p.DestX
          p.X = p.X + 6
        Case Is > p.DestX
            p.X = p.X - 6
        Case Is = p.DestX
            If isValue(p.dir, nDir.Left, nDir.Right) And p.Walking = True Then p.Walking = False: WalkOn GetTile
    End Select

    Select Case p.Y
        Case Is < p.DestY
            p.Y = p.Y + 6
        Case Is > p.DestY
            p.Y = p.Y - 6
        Case Is = p.DestY
            If isValue(p.dir, nDir.up, nDir.down) And p.Walking = True Then p.Walking = False: WalkOn GetTile
    End Select

Else

    p.WalkC = p.WalkC + 1

End If

If p.StepC > 5 Then

    p.StepC = 0
    p.Step = p.Step + 1: If p.Step > 1 Then p.Step = 0

Else

    p.StepC = p.StepC + 1

End If

If E.IsPressed(vbKeyUp) Then

    If p.Walking = False Then p.DestY = p.Y - 18: p.Walking = True: p.dir = up
    If T(GetTile).Walkable = False Then p.DestY = p.Y
    If isTileTaken(GetTile) = True Then p.DestY = p.Y
    
End If

If E.IsPressed(vbKeyDown) Then

    If p.Walking = False Then p.DestY = p.Y + 18: p.Walking = True: p.dir = down
    If T(GetTile).Walkable = False Then p.DestY = p.Y
    If isTileTaken(GetTile) = True Then p.DestY = p.Y
    
End If

If E.IsPressed(vbKeyLeft) Then

    If p.Walking = False Then p.DestX = p.X - 18: p.Walking = True: p.dir = Left
    If T(GetTile).Walkable = False Then p.DestX = p.X
    If isTileTaken(GetTile) = True Then p.DestX = p.X
    
End If

If E.IsPressed(vbKeyRight) Then

    If p.Walking = False Then p.DestX = p.X + 18: p.Walking = True: p.dir = Right
    If T(GetTile).Walkable = False Then p.DestX = p.X
    If isTileTaken(GetTile) = True Then p.DestX = p.X
    
End If

If E.IsPressed(vbKeySpace) Then TalkTo

If E.IsPressed(vbKeyZ) Then DoSpell
End Function

Function isTileTaken(I As Long)
Dim n
isTileTaken = False

        For n = 1 To 20
        
            If o(n).Act = True And o(n).DestX = T(I).X And o(n).DestY = T(I).Y Then isTileTaken = True: Exit Function
        
        Next n

End Function

Function LoadMarket()
Dim I As Long

frmGame.board.Visible = False
frmGame.menu.Visible = False
frmGame.battle.Visible = False
frmGame.main.Visible = False
frmGame.market.Visible = True

frmGame.lblMoney.Caption = p.Pack.money & "$"

frmGame.lstStuff.Clear
frmGame.lstICost.Clear

For I = 1 To 20

    If p.Pack.Item(I).Act = True Then
    
        frmGame.lstStuff.AddItem p.Pack.Item(I).Caption
        frmGame.lstICost.AddItem p.Pack.Item(I).Value
        
    End If

Next I

Dim n As String, n2

Open App.path + "\other data\" & M.MarketPath For Input As #1

Input #1, n

frmGame.market.Picture = LoadPicture(App.path + n)

frmGame.lstBuy.Clear
frmGame.lstMCost.Clear

Do Until EOF(1)

    Input #1, n, n2

    frmGame.lstBuy.AddItem n
    frmGame.lstMCost.AddItem n2

    DoEvents
    
Loop

Close #1
End Function

Function AddPackItem(Caption As String, Cost As Long) As Boolean
Dim I As Long

AddPackItem = True

For I = 1 To 20

    If p.Pack.Item(I).Act = False Then

        p.Pack.Item(I).Act = True
        p.Pack.Item(I).Caption = Caption
        p.Pack.Item(I).Value = Cost
        GoTo Done

    End If

Next I

AddPackItem = False
MsgBox ("Not enough room in your pack for the item!"), , "MagicQuest"

Done:
End Function

Function MoveObjs()
Dim I As Long

DoAI

For I = 1 To 20

       MoveObj o(I)
       
 Next I
End Function

Function MoveObj(p As obj)
Dim I As Long

If p.Walking = True And p.WalkC > 3 Then
    p.WalkC = 0

    Select Case p.X
        Case Is < p.DestX
          p.X = p.X + 6
        Case Is > p.DestX
            p.X = p.X - 6
        Case Is = p.DestX
            If isValue(p.dir, nDir.Left, nDir.Right) And p.Walking = True Then p.Walking = False: WalkOn GetTile
    End Select

    Select Case p.Y
        Case Is < p.DestY
            p.Y = p.Y + 6
        Case Is > p.DestY
            p.Y = p.Y - 6
        Case Is = p.DestY
            If isValue(p.dir, nDir.up, nDir.down) And p.Walking = True Then p.Walking = False: WalkOn GetTile
    End Select

Else

    p.WalkC = p.WalkC + 1

End If

If p.Person = True Then
If p.StepC > 5 Then

    p.StepC = 0
    p.Step = p.Step + 1: If p.Step > 1 Then p.Step = 0

Else

    p.StepC = p.StepC + 1

End If
End If

End Function

Function GetObj(Ind) As Long
Dim I As Long

For I = 1 To 20

    If o(I).Act = True And o(I).DestX = T(Ind).X And o(I).DestY = T(Ind).Y Then GetObj = I: Exit For

Next I
End Function

Function isPlayerOnTile(TileI As Long) As Boolean
isPlayerOnTile = False
If p.DestX = T(TileI).X And p.DestY = T(TileI).Y Then isPlayerOnTile = True
End Function

Function DoSpell()
Dim TileI As Long

Select Case p.dir

    Case nDir.up: TileI = GetTile - M.Mapheight
    Case nDir.down: TileI = GetTile + M.Mapheight
    Case nDir.Left: TileI = GetTile - 1
    Case nDir.Right: TileI = GetTile + 1

End Select

If T(TileI).Type = 3 And p.SpellCut = True And p.Pack.mDust > 0 Then

    p.Pack.mDust = p.Pack.mDust - 5: If p.Pack.mDust < 0 Then p.Pack.mDust = 0
    T(TileI).Type = 1
    T(TileI).Walkable = True
    
End If

If T(TileI).Type = 2 And p.SpellFreeze = True And p.Pack.mDust > 0 Then

    p.Pack.mDust = p.Pack.mDust - 5: If p.Pack.mDust < 0 Then p.Pack.mDust = 0
    T(TileI).Type = 15
    T(TileI).Walkable = True
    
End If
End Function

Function DoAI()
Dim I As Long

For I = 1 To 20

    If o(I).Act = True Then

        o(I).AI = o(I).AI + 1

        If o(I).AI > 200 And o(I).AI_Type = 1 Then
        
        Select Case o(I).dir
            Case nDir.up: o(I).dir = nDir.Right
            Case nDir.Right: o(I).dir = nDir.down
            Case nDir.down: o(I).dir = nDir.Left
            Case nDir.Left: o(I).dir = nDir.up
        End Select
        
        o(I).AI = 0
    
    Else
    
        o(I).AI = o(I).AI + 1

    End If

    If o(I).AI > 70 And o(I).AI_Type = 3 Then
    
        o(I).AI = 0
            
        Dim n As Long
        n = Int(Rnd * 8)

        Select Case n
            Case 4
                If o(I).Walking = False Then o(I).DestY = o(I).Y - 18: o(I).Walking = True: o(I).dir = nDir.up
                If T(GetObjTile(I)).Walkable = False Then o(I).DestY = o(I).Y: o(I).Walking = False
                If isTileTaken(GetObjTile(I)) = True Then o(I).DestY = o(I).Y: o(I).Walking = False
                If isPlayerOnTile(GetObjTile(I)) = True Then o(I).DestY = o(I).Y: o(I).Walking = False
            Case 5
                If o(I).Walking = False Then o(I).DestY = o(I).Y + 18: o(I).Walking = True: o(I).dir = nDir.down
                If T(GetObjTile(I)).Walkable = False Then o(I).DestY = o(I).Y: o(I).Walking = False
                If isTileTaken(GetObjTile(I)) = True Then o(I).DestY = o(I).Y: o(I).Walking = False
                If isPlayerOnTile(GetObjTile(I)) = True Then o(I).DestY = o(I).Y: o(I).Walking = False
            Case 6
                If o(I).Walking = False Then o(I).DestX = o(I).X - 18: o(I).Walking = True: o(I).dir = nDir.Left
                If T(GetObjTile(I)).Walkable = False Then o(I).DestX = o(I).X: o(I).Walking = False
                If isTileTaken(GetObjTile(I)) = True Then o(I).DestX = o(I).X: o(I).Walking = False
                If isPlayerOnTile(GetObjTile(I)) = True Then o(I).DestX = o(I).X: o(I).Walking = False
            Case 7
                If o(I).Walking = False Then o(I).DestX = o(I).X + 18: o(I).Walking = True: o(I).dir = nDir.Right
                If T(GetObjTile(I)).Walkable = False Then o(I).DestX = o(I).X: o(I).Walking = False
                If isTileTaken(GetObjTile(I)) = True Then o(I).DestX = o(I).X: o(I).Walking = False
                If isPlayerOnTile(GetObjTile(I)) = True Then o(I).DestX = o(I).X: o(I).Walking = False
            Case Else
                o(I).dir = n
            End Select

        Else
        
            o(I).AI = o(I).AI + 1
            
        End If
        
    End If

Next I
End Function


Function LoadSetBattle(Name, HP, Attack, Defense, pic As PictureBox)
frmGame.tmrAttack.Enabled = False
frmGame.tmrAttack2 = False
frmGame.tmrMagic.Enabled = False
frmGame.imgMagic.Visible = False
frmGame.imgAttack.Visible = False
frmGame.imgAttack2.Visible = False
frmGame.imgEnemy.Picture = pic.Picture
frmGame.board.Visible = False
frmGame.menu.Visible = False
frmGame.battle.Visible = True
frmGame.main.Visible = False
frmGame.market.Visible = False

Enem.Name = Name
Enem.HP = HP
Enem.MaxHP = Enem.HP
Enem.Attack = Attack
Enem.Defense = Defense
End Function

Function isNight()
isNight = False
Select Case Hour(Time)
Case Is < 6 Or Hour(Time) > 20: isNight = True
Case Is > 6 And Hour(Time) < 20: isNight = False
End Select
End Function

Function DoMsg(message As String)
    Dim money
    money = Int(Rnd * 20) + 10
    Msg.message = message
    Msg.Time = Len(Msg.message) + 10
End Function
