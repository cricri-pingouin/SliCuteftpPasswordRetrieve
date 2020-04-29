VERSION 5.00
Begin VB.Form frmDecrypt 
   Caption         =   "CuteFtp Decrypter"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   Icon            =   "frmDecrypt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Enter the full login name (case sensitive)"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdSeekPwd 
      Caption         =   "Seek"
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtSiteName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter the site name fully or partially (case sensitive)"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtDecrypted 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Decrypted password"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   855
      Left            =   2040
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Encrypted password"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLogin 
      Caption         =   "Exact login name:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblSiteName 
      Caption         =   "Site name:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblDecrypted 
      Caption         =   "Decrypted password:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblPassword 
      Caption         =   "Encrypted password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frmDecrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Original As String

Private Sub cmdAbout_Click()
   Original = "CuteFtp Password Decrypter v1.0" + Chr(13)
   Original = Original + "Two possible uses:" + Chr(13) + "* Enter part of a site name in 'Site name', the full login in 'Exact login name',"
   Original = Original + Chr(13) + "both case sensitive, and press 'Seek'."
   Original = Original + Chr(13) + "* Open your 'SM.DAT' with an editor, find the encrypted password,"
   Original = Original + Chr(13) + "copy and paste it in 'Encrypted password' and press 'Decrypt'."
   Original = Original + Chr(13) + Chr(13) + "You need SM.DAT in the program directory, and if everything is fine,"
   Original = Original + Chr(13) + "the decrypted password will appear in 'Decrypted password'."
   MsgBox Original, vbInformation, "About this program"
End Sub

Private Sub cmdDecrypt_Click()
   Dim i As Integer
   
   Original = txtPassword.Text
   If Original = "" Then
      MsgBox "You have to enter an encrypted password to proceed.", vbInformation, "Missing information"
      Exit Sub
   End If
   txtDecrypted.Text = ""
   For i = 1 To Len(Original)
      'For each character do "Character Xor #C8h"
      txtDecrypted = txtDecrypted + Chr(Asc(Mid(Original, i, 1)) Xor 200)
   Next
End Sub

Private Sub cmdSeekPwd_Click()
   Dim SiteNameFound As Boolean
   Dim ReadByte As String * 1
   Dim i, Position As Integer
   
   'Check that SM.DAT exists in program directory
   Original = Dir("sm.dat", 0)
   If Original = "" Then
      MsgBox "The file SM.DAT was not found in the program directory.", vbCritical, "File not found"
      Exit Sub
   End If
   'Check if a site name and login have been entered
   If txtSiteName.Text = "" Or txtLogin.Text = "" Then
      MsgBox "You have to enter a site name and a login to proceed.", vbExclamation, "Missing information"
      Exit Sub
   End If
   Original = txtSiteName.Text
   SiteNameFound = False
   Open "sm.dat" For Binary Access Read As #1
   'Search site name
   Do
      Get #1, , ReadByte
      If ReadByte = Left(Original, 1) Then
         'First letter found
         SiteNameFound = True
         'Check next characters
         For i = 2 To Len(Original)
            Get #1, , ReadByte
            If ReadByte <> Mid(Original, i, 1) Then
               SiteNameFound = False
               Exit For
            End If
         Next
      End If
   Loop Until EOF(1) Or SiteNameFound
   'Search full login name
   If SiteNameFound Then
      Original = txtLogin.Text
      SiteNameFound = False
      Do
         Get #1, , ReadByte
         If ReadByte = Left(Original, 1) Then
            'First letter found
            SiteNameFound = True
            'Check next characters
            For i = 2 To Len(Original)
               Get #1, , ReadByte
               If ReadByte <> Mid(Original, i, 1) Then
                  SiteNameFound = False
                  Exit For
               End If
            Next
         End If
      Loop Until EOF(1) Or SiteNameFound
      'Grab encrypted password string
      If SiteNameFound Then
         'Next character is password length
         Get #1, , ReadByte
         txtPassword = ""
         'Grab encrypted password
         For i = 1 To Asc(ReadByte)
            Get #1, , ReadByte
            txtPassword = txtPassword + ReadByte
         Next
         'Decrypt password
         Call cmdDecrypt_Click
      Else
         MsgBox "A site was found but the login does not match.", vbExclamation, "Wrong login"
      End If
   Else
      Original = "No site was found with '" + txtSiteName.Text + "' as part of the name."
      MsgBox Original, vbExclamation, "Site name not found"
   End If
   Close #1
End Sub
