VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "World Generator"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   4230
      Left            =   5025
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   324
      TabIndex        =   8
      Top             =   -15
      Width           =   4920
   End
   Begin VB.TextBox txtIterations 
      Height          =   285
      Left            =   3015
      TabIndex        =   7
      Text            =   "500"
      Top             =   4980
      Width           =   615
   End
   Begin VB.TextBox txtWater 
      Height          =   285
      Left            =   3015
      TabIndex        =   5
      Text            =   "30"
      Top             =   4305
      Width           =   615
   End
   Begin VB.TextBox txtIce 
      Height          =   285
      Left            =   3015
      TabIndex        =   4
      Text            =   "3"
      Top             =   4635
      Width           =   615
   End
   Begin VB.CommandButton cmdGenereer 
      Caption         =   "Generate"
      Height          =   660
      Left            =   60
      TabIndex        =   1
      Top             =   4290
      Width           =   1950
   End
   Begin VB.PictureBox Picture1 
      Height          =   4230
      Left            =   30
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   324
      TabIndex        =   0
      Top             =   0
      Width           =   4920
   End
   Begin VB.Label lblPersistent 
      Caption         =   "http://www.persistentrealities.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   3780
      MousePointer    =   2  'Cross
      TabIndex        =   10
      Top             =   5175
      Width           =   2565
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":08CA
      Height          =   465
      Left            =   3750
      TabIndex        =   9
      Top             =   4290
      Width           =   6225
   End
   Begin VB.Label lblIterations 
      Caption         =   "Iterations:"
      Height          =   240
      Left            =   2250
      TabIndex        =   6
      Top             =   5025
      Width           =   720
   End
   Begin VB.Label lblIce 
      Caption         =   "% Ice"
      Height          =   270
      Left            =   2430
      TabIndex        =   3
      Top             =   4650
      Width           =   540
   End
   Begin VB.Label lblWater 
      Caption         =   "% Water:"
      Height          =   180
      Left            =   2190
      TabIndex        =   2
      Top             =   4290
      Width           =   780
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//World Generator
'//Ported from http://www.lysator.liu.se/~johol/fwmg/fwmg.html.
'//Date: 9 November 20002
'//Version: 1
'//by Almar Joling, http://www.persistentrealities.com


Private Const PI As Single = 3.141593
Private Const Max_Rand As Long = 2147483647

Dim WorldMapArray() As Integer
Private WorldMapArraySphere() As Integer

Private XRange As Long
Private YRange As Long
Private Histogram() As Long
Private FilledPixels As Integer
Private Red, Green, Blue
Private YRangeDiv2  As Single, YRangeDivPI As Single
Private SinIterPhi() As Single
Private Int_Min As Integer
Private Seed As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Init()
    XRange = 256 '320
    YRange = 256
    Red = Array(0, 0, 0, 0, 0, 0, 0, 0, 34, 68, 102, 119, 136, 153, 170, 187, 0, 34, 34, 119, 187, 255, 238, 221, 204, 187, 170, 153, 136, 119, 85, 68, 255, 250, 245, 240, 235, 230, 225, 220, 215, 210, 205, 200, 195, 190, 185, 180, 175)
    Green = Array(0, 0, 17, 51, 85, 119, 153, 204, 221, 238, 255, 255, 255, 255, 255, 255, 68, 102, 136, 170, 221, 187, 170, 136, 136, 102, 85, 85, 68, 51, 51, 34, 255, 250, 245, 240, 235, 230, 225, 220, 215, 210, 205, 200, 195, 190, 185, 180, 175)
    Blue = Array(0, 68, 102, 136, 170, 187, 221, 255, 255, 255, 255, 255, 255, 255, 255, 255, 0, 0, 0, 0, 0, 34, 34, 34, 34, 34, 34, 34, 34, 34, 17, 0, 255, 250, 245, 240, 235, 230, 225, 220, 215, 210, 205, 200, 195, 190, 185, 180, 175)
    Int_Min = -256
End Sub

Private Sub FloodFill4(X As Long, Y As Long, OldColor As Integer)
    If WorldMapArray(X * YRange + Y) = OldColor Then
        If (WorldMapArray(X * YRange + Y) < 16) Then
            WorldMapArray(X * YRange + Y) = 32
        Else
            WorldMapArray(X * YRange + Y) = WorldMapArray(X * YRange + Y) + 17
        End If
        
        FilledPixels = FilledPixels + 1
        
        If (Y - 1 > 0) Then Call FloodFill4(X, Y - 1, OldColor)
        If (Y + 1 < YRange) Then Call FloodFill4(X, Y + 1, OldColor)
        
        If (X - 1 < 0) Then
            Call FloodFill4(XRange - 1, Y, OldColor)
        Else
            Call FloodFill4(X - 1, Y, OldColor)
        End If
    
        If (X + 1 > XRange - 1) Then
            Call FloodFill4(0, Y, OldColor)
        Else
            Call FloodFill4(X + 1, Y, OldColor)
        End If
    End If
End Sub

Private Sub GenerateWorldMap()
    Dim alpha As Single, Beta As Single
    Dim TanB As Single
    Dim I As Integer, Row As Long, N2 As Integer
    Dim Theta As Long, Phi As Integer, Xsi As Integer
    Dim Flag1 As Integer
    
    Flag1 = Rand And 1
    Row = 0
    
    alpha = (Rand / Max_Rand - 0.5) * PI '//Rotate around x-axis
    Beta = (Rand / Max_Rand - 0.5) * PI  '//Rotate around y-axis
    TanB = Tan(ACos(Cos(alpha) * Cos(Beta)))
    Xsi = (XRange / 2 - (XRange / PI) * Beta)
    
    For Phi = 0 To (XRange / 2) - 1 '<---
        Theta = (YRangeDivPI * Atan(SinIterPhi(Xsi - Phi + XRange) * TanB)) + YRangeDiv2
        
        If (Flag1) Then
            '//Rise northen hemisphere <=> lower southern
            If (WorldMapArray(Row + Theta) <> Int_Min) Then
                WorldMapArray(Row + Theta) = WorldMapArray(Row + Theta) - 1
            Else
                WorldMapArray(Row + Theta) = -1
            End If
    
        Else
            '//Rise southern hemisphere
            If (WorldMapArray(Row + Theta) <> Int_Min) Then
                WorldMapArray(Row + Theta) = WorldMapArray(Row + Theta) + 1
            Else
                WorldMapArray(Row + Theta) = 1
            End If
        End If
        
        Row = Row + YRange
    Next Phi
End Sub


Private Sub Main(lngDC As Long, Optional blnGlobal As Boolean = False)
    Dim NumberOfFaults As Long
    Dim A As Long, J As Long, I As Long, Color As Integer
    Dim MaxZ As Integer, MinZ As Integer
    Dim Row As Long, TwoColorMode As Integer
    Dim Index2 As Long
    Dim Threshold As Long, Count As Long
    Dim PercentWater As Integer, PercentIce As Integer, Cur As Integer
    
    '//Set options
    FilledPixels = 0
    NumberOfFaults = txtIterations.Text
    MaxZ = 10
    MinZ = -1
    PercentWater = txtWater.Text
    PercentIce = txtIce.Text
    TwoColorMode = False
    
    
    '//Clear all
    ReDim Histogram(256)
    ReDim WorldMapArray(XRange& * YRange&)
    ReDim SinIterPhi(2 * XRange)
    
    For I = 0 To XRange - 1
        SinIterPhi(I) = Sin(I * 2 * PI / XRange)
        SinIterPhi(I + XRange) = Sin(I * 2 * PI / XRange)
    Next I
    
    Randomize -Seed
    
    Row = 0
    For J = 0 To XRange - 1
        WorldMapArray(Row) = 0
        
        For I = 1 To YRange - 1
            WorldMapArray(I + Row) = Int_Min
        Next I
        
        Row = Row + YRange
    Next J

    '//Define some "constants" which we use frequently
    YRangeDiv2 = YRange / 2
    YRangeDivPI = YRange / PI

    '//Generate the map!
    For A = 0 To NumberOfFaults - 1
        GenerateWorldMap
    Next A


    '//Copy data (I have only calculated faults for 1/2 the image.
    '//I can do this due to symmetry... :)
    Index2 = (XRange / 2) * YRange
    Row = 0
    For J = 0 To (XRange / 2) - 1
        For I = 0 To YRange - 1
            WorldMapArray(Row + Index2 + YRange - I) = WorldMapArray(Row + I)
        Next I
     
        Row = Row + YRange
    Next J
  

    '//Reconstruct the real WorldMap from the WorldMapArray and FaultArray
    Row = 0
    For J = 0 To XRange - 1

        '//We have to start somewhere, and the top row was initialized to 0,
        '//but it might have changed during the iterations...
        Color = WorldMapArray(Row)
    
        For I = 1 To YRange - 1
            '// We "fill" all positions with values != INT_MIN with Color
            Cur = WorldMapArray(Row + I)
            If (Cur <> Int_Min) Then
                Color = Color + Cur
            End If
    
            WorldMapArray(Row + I) = Color
        Next I
    
        Row = Row + YRange
    Next J
 

    '//Compute MAX and MIN values in WorldMapArray
    For J = 0 To (XRange * YRange) - 1
        Color = WorldMapArray(J)
        
        If (Color > MaxZ) Then MaxZ = Color
        If (Color < MinZ) Then MinZ = Color
    Next J


    '//Compute color-histogram of WorldMapArray.
    '//This histogram is a very crude aproximation, since all pixels are
    '//considered of the same size... I will try to change this in a
    '//later version of this program.
    Row = 0
    For J = 0 To XRange - 1
        For I = 0 To YRange - 1
            Color = WorldMapArray(Row + I)
            Color = ((Color - MinZ + 1) / (MaxZ - MinZ + 1)) * 30 + 1
            Histogram(Color) = Histogram(Color) + 1
        Next I
        
        Row = Row + YRange
    Next J


    '//Threshold now holds how many pixels PercentWater means
    Threshold = (PercentWater * XRange * YRange) / 100

    
    '//"Integrate" the histogram to decide where to put sea-level
    Count = 0
    For J = 0 To 256 - 1
        Count = Count + Histogram(J)
        If (Count > Threshold) Then Exit For
    Next J


    '//Threshold now holds where sea-level is
    Threshold = J * (MaxZ - MinZ + 1) / 30 + MinZ

    If (TwoColorMode) Then
        Row = 0
        For J = 0 To XRange - 1
            For I = 0 To YRange - 1
                Color = WorldMapArray(Row + I)
    
                If (Color < Threshold) Then
                    WorldMapArray(Row + I) = 3
                Else
                    WorldMapArray(Row + I) = 20
                End If
            Next I
            Row = Row + YRange
        Next J

  Else

    '//Scale WorldMapArray to colorrange in a way that gives you
    '//a certain Ocean/Land ratio
     
    Row = 0
    For J = 0 To XRange - 1
        For I = 0 To YRange - 1
    
            Color = WorldMapArray(Row + I)
        
            If (Color < Threshold) Then
                Color = ((Color - MinZ) / (Threshold - MinZ) * 15) + 1
            Else
                Color = (((Color - Threshold) / (MaxZ - Threshold)) * 15) + 16
            End If
        
            '// Just in case... I DON't want the GIF-saver to flip out! :)
            If (Color < 1) Then Color = 1
            If (Color > 255) Then Color = 31
            WorldMapArray(Row + I) = Color
    
        Next I
        Row = Row + YRange
    Next J


    '// "Recycle" Threshold variable, and, eh, the variable still has something
    '// like the same meaning... :)
    Threshold = PercentIce * XRange * YRange / 100

    If ((Threshold <= 0) Or (Threshold > XRange * YRange)) Then GoTo Finished

    FilledPixels = 0

    '//i==y, j=x
    
    For I = 0 To YRange - 1
        Row = 0
        For J = 0 To XRange - 1
    
            Color = WorldMapArray(Row + I)
            If (Color < 32) Then Call FloodFill4(J, I, Color)
    
            '//FilledPixels is a global variable which FloodFill4 modifies...
            '//I know it's ugly, but as it is now, this is a hack! :)
            If (FilledPixels > Threshold) Then GoTo NorthPoleFinished
            
            Row = Row + YRange
        Next J
    Next I
  End If
  
NorthPoleFinished:
    FilledPixels = 0

    '//i==y, j==x
    For I = YRange To 0 Step -1
        Row = 0
        For J = 0 To XRange - 1
            
            Color = WorldMapArray(Row + I)
            If (Color < 32) Then Call FloodFill4(J, I, Color)
        
        
            '//FilledPixels is a global variable which FloodFill4 modifies...
            '//I know it's ugly, but as it is now, this is a hack! :)
            If (FilledPixels > Threshold) Then GoTo Finished
            Row = Row + YRange
        Next J
    Next I
   
Finished:

    Smoothning
    
    '//Global or not?
    If blnGlobal = True Then
        MakeGlobal 0
    Else
        WorldMapArraySphere = WorldMapArray
    End If
    
    '// i==y, j==x
    For I = 0 To YRange - 1
        Row = 0
        For J = 0 To XRange - 1
            Color = WorldMapArraySphere(Row + I)
            SetPixel lngDC, J, I, RGB(Red(Color), Green(Color), Blue(Color))
            Row = Row + YRange
        Next J
    Next I
End Sub


Private Sub cmdGenereer_Click()
    Seed = Rnd * 32768
    Init
    
    XRange = 320: YRange = 160
    Call Main(Picture1.hdc)
    
    XRange = 256: YRange = 256
    Call Main(Picture2.hdc, True)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub

Private Sub MakeGlobal(sngAngle As Single)
    Dim sngDestH As Long, sngDestW As Long
    Dim fx As Single, fy As Single, tx As Single, ty As Single
    Dim Y As Long, X As Long, W As Single
    Dim alpha As Single, q As Single
    Dim Offs As Single
    'Dim sngAngle As Single
    Dim Row As Long
    
    Dim ArrayTemp() As Integer
    ReDim ArrayTemp(XRange * YRange)
    
    sngAngle = 80
    Offs = (sngAngle * XRange) / 360
    
    sngDestH = YRange \ 2
    sngDestW = XRange \ 2
   
    For Y = 0 To YRange - 1 Step 1
        q = 1! * Abs(sngDestH - Y) / sngDestH
        W = sngDestW * Sqr(1! - q! * q!)
        Row = 0
        For X = -W + 1 To W - 1  ' Step 1
            ty = Y
            tx = sngDestW + X
            fy = Y
            alpha = ACos2((1! * X) / W) / PI
            
            fx = (Offs! + alpha! * sngDestW + 1! * XRange) Mod XRange
            ArrayTemp(tx * YRange + ty) = WorldMapArray(fx * YRange + fy)
        Next X
        Row = Row + (YRange - 1)
    Next Y
    
    '//Copy new array over old one
    
    WorldMapArraySphere = ArrayTemp
End Sub


Private Function ACos2(Num As Single) As Single
    'On Error Resume Next
    '//Get the Acos
    ACos2 = Atn(-Num / Sqr(-Num * Num + 1)) + 1.5707963267949
End Function

Private Function Rand() As Double
    Rand = Rnd * Max_Rand
End Function


Public Function ACos#(Num As Double)
    ACos = Atn((Num * -1) / Sqr((Num * -1) * Num + 1)) + 2 * Atn(1)
End Function

Public Function Atan#(Num As Double)
    Atan = Atn(Num)
End Function


Private Sub Smoothning()
    Dim X As Long, Z As Long, Row As Long
    Dim K As Single
    K = 0.8
    
    '// Rows, left to right
    For Z = 1 To YRange - 1
        Row = 0
        For X = 0 To XRange - 1
            WorldMapArray(Row + Z) = WorldMapArray(Row + 1 + Z) * (1 - K) + WorldMapArray(Row + Z) * K
            Row = Row + YRange
        Next X
    Next Z

    '//Rows, right to left
    For Z = 1 To YRange - 1
        Row = 0
        For X = 0 To XRange - 1
            WorldMapArray(Row + Z) = WorldMapArray(Row - 1 + Z) * (1 - K) + WorldMapArray(Row + Z) * K
            Row = Row + YRange
        Next X
    Next Z
    

    '//Rows, right to left
    For Z = 1 To YRange - 1
        Row = 0
        For X = 0 To XRange - 1
            WorldMapArray(Row + Z) = WorldMapArray(Row + Z + 1) * (1 - K) + WorldMapArray(Row + Z) * K
            Row = Row + YRange
        Next X
    Next Z
    
    
    '//Rows, right to left
    For Z = 1 To YRange - 1
        Row = 0
        For X = 0 To XRange - 1
            WorldMapArray(Row + Z) = WorldMapArray(Row + Z - 1) * (1 - K) + WorldMapArray(Row + Z) * K
            Row = Row + YRange
        Next X
    Next Z

End Sub

Private Sub lblPersistent_Click()
    Call ShellExecute(0, "OPEN", "http://www.persistentrealities.com", 0, App.Path, 1)
End Sub
