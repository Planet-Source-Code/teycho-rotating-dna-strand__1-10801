VERSION 5.00
Begin VB.Form frmDNA 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotating DNA Strand"
   ClientHeight    =   8865
   ClientLeft      =   5025
   ClientTop       =   2640
   ClientWidth     =   4530
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   591
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   302
End
Attribute VB_Name = "frmDNA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NumOfBars = 20
Private DNA As DNAStrand

Private Type DNAMolecule
   X As Integer
   LastX As Integer
   Radius As Integer
   LastRad As Integer
   BorderColor As Long
   FillColor As Long
End Type

Private Type DNABars
   Angle As Single
   Y As Integer
   BarColor As Long
   Mol1 As DNAMolecule
   Mol2 As DNAMolecule
End Type

Private Type DNAProperties
   Center As Integer
   BarLength As Integer
   MolRadius As Integer
   OBS As Integer
   Radius As Integer
End Type

Private Type DNAStrand
   Properties As DNAProperties
   Bars(NumOfBars) As DNABars
End Type

Public jInterval

Sub timeout(ms)

StartTime = Timer
Do
Loop Until Timer > StartTime + ms

End Sub

Private Sub Form_Activate()

For QAngle = 0 To 32000 Step jInterval
   For Bar = 1 To NumOfBars
      DoEvents
      
      jAngle = QAngle + ((Bar * 0.5) + 0.5)
      
      DNA.Bars(Bar).Mol1.X = (DNA.Properties.Center - (Cos(jAngle) * (DNA.Properties.BarLength / 2)))
      DNA.Bars(Bar).Mol2.X = (DNA.Properties.Center + (Cos(jAngle) * (DNA.Properties.BarLength / 2)))
      
      a1 = DNA.Properties.OBS - (Sin(jAngle) * (DNA.Properties.BarLength))
      a2 = DNA.Properties.OBS + (Sin(jAngle) * (DNA.Properties.BarLength))
    
      DNA.Bars(Bar).Mol1.Radius = (DNA.Properties.OBS / a1) * DNA.Properties.Radius
      DNA.Bars(Bar).Mol2.Radius = (DNA.Properties.OBS / a2) * DNA.Properties.Radius
      
      Me.FillColor = &H0
      Me.Line (DNA.Bars(Bar).Mol1.LastX, DNA.Bars(Bar).Y)-(DNA.Bars(Bar).Mol2.LastX, DNA.Bars(Bar).Y), 0
      Me.Circle (DNA.Bars(Bar).Mol1.LastX, DNA.Bars(Bar).Y), DNA.Bars(Bar).Mol1.LastRad, 0
      Me.Circle (DNA.Bars(Bar).Mol2.LastX, DNA.Bars(Bar).Y), DNA.Bars(Bar).Mol2.LastRad, 0

      Me.Line (DNA.Bars(Bar).Mol1.X, DNA.Bars(Bar).Y)-(DNA.Bars(Bar).Mol2.X, DNA.Bars(Bar).Y), DNA.Bars(Bar).BarColor
      If DNA.Bars(Bar).Mol1.Radius > DNA.Bars(Bar).Mol2.Radius Then
         Me.FillColor = DNA.Bars(Bar).Mol2.FillColor
         Me.Circle (DNA.Bars(Bar).Mol2.X, DNA.Bars(Bar).Y), DNA.Bars(Bar).Mol2.Radius, DNA.Bars(Bar).Mol1.BorderColor
         Me.FillColor = DNA.Bars(Bar).Mol1.FillColor
         Me.Circle (DNA.Bars(Bar).Mol1.X, DNA.Bars(Bar).Y), DNA.Bars(Bar).Mol1.Radius, DNA.Bars(Bar).Mol1.BorderColor
      Else
         Me.FillColor = DNA.Bars(Bar).Mol1.FillColor
         Me.Circle (DNA.Bars(Bar).Mol1.X, DNA.Bars(Bar).Y), DNA.Bars(Bar).Mol1.Radius, DNA.Bars(Bar).Mol1.BorderColor
         Me.FillColor = DNA.Bars(Bar).Mol2.FillColor
         Me.Circle (DNA.Bars(Bar).Mol2.X, DNA.Bars(Bar).Y), DNA.Bars(Bar).Mol2.Radius, DNA.Bars(Bar).Mol1.BorderColor
      End If
      DNA.Bars(Bar).Mol1.LastX = DNA.Bars(Bar).Mol1.X
      DNA.Bars(Bar).Mol2.LastX = DNA.Bars(Bar).Mol2.X
      
      DNA.Bars(Bar).Mol1.LastRad = DNA.Bars(Bar).Mol1.Radius
      DNA.Bars(Bar).Mol2.LastRad = DNA.Bars(Bar).Mol2.Radius
      
   Next Bar
   jInterval = jInterval + 0.01
   timeout 0.01
Next QAngle

End Sub

Private Sub Form_Click()
MsgBox Me.Width
MsgBox Me.Height
End Sub

Private Sub Form_Load()

For q = 1 To NumOfBars
   DNA.Bars(q).Y = (q * 25) + 25
Next
jInterval = 0.3
DNA.Properties.Center = 150
DNA.Properties.BarLength = 100
DNA.Properties.Radius = 4.5
DNA.Properties.OBS = 175
For q = 1 To NumOfBars
   With DNA.Bars(q)
      .BarColor = &H808080
      .Mol1.BorderColor = &HFFFFFF
      .Mol2.BorderColor = &HFFFFFF
      .Mol1.FillColor = &HFF0000
      .Mol2.FillColor = &HFF
   End With
Next q
End Sub


