VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDemoAppMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CLASSter Demo Application"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCustIDs 
      Height          =   315
      ItemData        =   "frmDemoAppMain.frx":0000
      Left            =   1845
      List            =   "frmDemoAppMain.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   75
      Width           =   1620
   End
   Begin VB.CommandButton cmdNextPage 
      Caption         =   "Next"
      Height          =   330
      Left            =   8130
      TabIndex        =   3
      Top             =   4905
      Width           =   1200
   End
   Begin VB.CheckBox chkStopInProc 
      Caption         =   "Check this box to step into code"
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   4980
      Width           =   3375
   End
   Begin VB.CommandButton cmdPrevPage 
      Caption         =   "Previous"
      Height          =   330
      Left            =   6840
      TabIndex        =   1
      Top             =   4920
      Width           =   1200
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   4320
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   7620
      _Version        =   393216
      Tabs            =   12
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "1. Connection"
      TabPicture(0)   =   "frmDemoAppMain.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "2. Simple SQLs"
      TabPicture(1)   =   "frmDemoAppMain.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "3. SQL, Emb. Prms"
      TabPicture(2)   =   "frmDemoAppMain.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "4. Parameterized SQLs"
      TabPicture(3)   =   "frmDemoAppMain.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "5. Calling SPs"
      TabPicture(4)   =   "frmDemoAppMain.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "6. ExecSPbyName"
      TabPicture(5)   =   "frmDemoAppMain.frx":0090
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "7. Transactions"
      TabPicture(6)   =   "frmDemoAppMain.frx":00AC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame4"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "8. Error Handling"
      TabPicture(7)   =   "frmDemoAppMain.frx":00C8
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "frames(3)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "9. Batch Execution"
      TabPicture(8)   =   "frmDemoAppMain.frx":00E4
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame3"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "10. Auto-Formatting"
      TabPicture(9)   =   "frmDemoAppMain.frx":0100
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame10"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "11. XML, Basic Level"
      TabPicture(10)  =   "frmDemoAppMain.frx":011C
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame2"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "12. XML, Advanced"
      TabPicture(11)  =   "frmDemoAppMain.frx":0138
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame11"
      Tab(11).ControlCount=   1
      Begin VB.Frame Frame11 
         Caption         =   "12. XML Functionality, Advanced Level"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   64
         Top             =   1020
         Width           =   8940
         Begin VB.TextBox txtNodeSpecs 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   375
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            Top             =   720
            Width           =   8400
         End
         Begin VB.CommandButton cmdBuildXMLBasic2 
            Caption         =   "Construct XML"
            Height          =   360
            Left            =   6780
            TabIndex        =   68
            Top             =   1665
            Width           =   2025
         End
         Begin VB.TextBox txtSkipRecs 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1890
            TabIndex        =   67
            Text            =   "2"
            Top             =   1650
            Width           =   510
         End
         Begin VB.TextBox txtIncRecs 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4290
            TabIndex        =   66
            Text            =   "4"
            Top             =   1620
            Width           =   510
         End
         Begin VB.ComboBox cboNodeSpecs 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmDemoAppMain.frx":0154
            Left            =   2115
            List            =   "frmDemoAppMain.frx":0161
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   315
            Width           =   6690
         End
         Begin VB.Label Label32 
            Caption         =   $"frmDemoAppMain.frx":02F1
            ForeColor       =   &H00FF0000&
            Height          =   645
            Left            =   195
            TabIndex        =   83
            Top             =   2310
            Width           =   7935
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Nodes Specification:"
            Height          =   270
            Left            =   240
            TabIndex        =   71
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Records to skip:"
            Height          =   300
            Left            =   375
            TabIndex        =   70
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Records to include:"
            Height          =   300
            Left            =   2655
            TabIndex        =   69
            Top             =   1665
            Width           =   1515
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "11. XML Functionality, Basic Level"
         ForeColor       =   &H00FF0000&
         Height          =   3120
         Left            =   -74850
         TabIndex        =   53
         Top             =   1020
         Width           =   8940
         Begin VB.ComboBox cboOpenTag 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmDemoAppMain.frx":03D0
            Left            =   375
            List            =   "frmDemoAppMain.frx":03DA
            TabIndex        =   62
            Text            =   "CustInfo"
            Top             =   540
            Width           =   8385
         End
         Begin VB.TextBox txtCloseTag 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2055
            TabIndex        =   60
            Top             =   945
            Width           =   2790
         End
         Begin VB.TextBox txtRecElems 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2085
            TabIndex        =   56
            Text            =   "Customer,Order"
            Top             =   1620
            Width           =   2760
         End
         Begin VB.TextBox txtRSElems 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2085
            TabIndex        =   55
            Text            =   ",Orders"
            Top             =   1275
            Width           =   2730
         End
         Begin VB.CommandButton cmdBuildXMLBasic 
            Caption         =   "Construct XML"
            Height          =   330
            Left            =   6630
            TabIndex        =   54
            Top             =   1590
            Width           =   2100
         End
         Begin VB.Label Label31 
            Caption         =   $"frmDemoAppMain.frx":0430
            ForeColor       =   &H00FF0000&
            Height          =   765
            Left            =   270
            TabIndex        =   82
            Top             =   2265
            Width           =   8475
         End
         Begin VB.Label Label18 
            Caption         =   "(Optional) Closing Tag:"
            Height          =   300
            Left            =   360
            TabIndex        =   61
            Top             =   915
            Width           =   1830
         End
         Begin VB.Label Label17 
            Caption         =   "zRow Elements:"
            Height          =   300
            Left            =   420
            TabIndex        =   59
            Top             =   1680
            Width           =   1485
         End
         Begin VB.Label Label6 
            Caption         =   "RsData Elements:"
            Height          =   300
            Left            =   375
            TabIndex        =   58
            Top             =   1305
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Root Element/Opening tag:"
            Height          =   270
            Left            =   360
            TabIndex        =   57
            Top             =   285
            Width           =   2070
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "10. Automatic Fields Formatting"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   43
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdShowOrderFmtted 
            Caption         =   "Execute"
            Height          =   330
            Left            =   7080
            TabIndex        =   47
            Top             =   1710
            Width           =   1710
         End
         Begin VB.TextBox txtFmtMoney 
            Height          =   300
            Left            =   1905
            TabIndex        =   46
            Text            =   "$#,##0.00"
            Top             =   345
            Width           =   1530
         End
         Begin VB.TextBox txtFmtDateDft 
            Height          =   300
            Left            =   1920
            TabIndex        =   45
            Text            =   "MM/DD/YYYY"
            Top             =   660
            Width           =   1530
         End
         Begin VB.TextBox txtFmtShipDate 
            Height          =   300
            Left            =   1920
            TabIndex        =   44
            Text            =   "MMM DD, YYYY"
            Top             =   990
            Width           =   1530
         End
         Begin VB.Label Label30 
            Caption         =   $"frmDemoAppMain.frx":0564
            ForeColor       =   &H00FF0000&
            Height          =   765
            Left            =   210
            TabIndex        =   81
            Top             =   2220
            Width           =   7860
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Money Format:"
            Height          =   300
            Left            =   570
            TabIndex        =   52
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Default Date Format:"
            Height          =   300
            Left            =   150
            TabIndex        =   51
            Top             =   705
            Width           =   1575
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Ship Date Format:"
            Height          =   300
            Left            =   165
            TabIndex        =   50
            Top             =   1020
            Width           =   1575
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Found Order:"
            Height          =   300
            Left            =   510
            TabIndex        =   49
            Top             =   1350
            Width           =   1200
         End
         Begin VB.Label lblFoundOrder 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Found Order:"
            Height          =   300
            Left            =   1935
            TabIndex        =   48
            Top             =   1320
            Width           =   6840
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "9. Batch Execution, Multiple Recordsets"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   39
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdExecBatch 
            Caption         =   "Execute Batch"
            Height          =   330
            Left            =   6495
            TabIndex        =   41
            Top             =   1575
            Width           =   2310
         End
         Begin VB.TextBox txtBatch 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   1875
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   405
            Width           =   6855
         End
         Begin VB.Label Label29 
            Caption         =   $"frmDemoAppMain.frx":066F
            ForeColor       =   &H00FF0000&
            Height          =   795
            Left            =   195
            TabIndex        =   80
            Top             =   2220
            Width           =   8220
         End
         Begin VB.Label Label1 
            Caption         =   "Executed batch:"
            Height          =   300
            Left            =   315
            TabIndex        =   42
            Top             =   420
            Width           =   1350
         End
      End
      Begin VB.Frame frames 
         Caption         =   "8. Error Handling"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Index           =   3
         Left            =   -74850
         TabIndex        =   30
         Top             =   1020
         Width           =   8940
         Begin VB.CheckBox chkErrAbort 
            Caption         =   "Abort transaction"
            Height          =   285
            Left            =   3810
            TabIndex        =   36
            Top             =   945
            Width           =   1575
         End
         Begin VB.CheckBox chkErrTrans 
            Caption         =   "Execute in transaction"
            Height          =   285
            Left            =   255
            TabIndex        =   35
            Top             =   960
            Width           =   2355
         End
         Begin VB.TextBox txtErrSQL 
            Height          =   300
            Left            =   255
            TabIndex        =   34
            Text            =   "SELECT x,y,z  FROM Customers"
            Top             =   555
            Width           =   5700
         End
         Begin VB.CommandButton cmdExecWithErrors1 
            Caption         =   "Execute Sample 1"
            Height          =   330
            Left            =   6840
            TabIndex        =   33
            Top             =   1155
            Width           =   1935
         End
         Begin VB.CheckBox chkErrRaise 
            Caption         =   "Reraise Error (Sample 1 Only)"
            Height          =   285
            Left            =   3825
            TabIndex        =   32
            Top             =   1260
            Width           =   2700
         End
         Begin VB.CommandButton cmdExecWithErrors2 
            Caption         =   "Execute Sample 2"
            Height          =   330
            Left            =   6795
            TabIndex        =   31
            Top             =   1590
            Width           =   1965
         End
         Begin VB.Label Label28 
            Caption         =   $"frmDemoAppMain.frx":07B3
            ForeColor       =   &H00FF0000&
            Height          =   930
            Left            =   195
            TabIndex        =   79
            Top             =   2115
            Width           =   8070
         End
         Begin VB.Label Label27 
            Caption         =   "ON ERROR:"
            Height          =   255
            Left            =   2730
            TabIndex        =   38
            Top             =   975
            Width           =   945
         End
         Begin VB.Label Label10 
            Caption         =   "Enter invalid SQL statement to generate database errors"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   225
            TabIndex        =   37
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "7. Using transactions"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   26
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdIncCustOrders 
            Caption         =   "Increment All Freights"
            Height          =   330
            Left            =   6420
            TabIndex        =   29
            Top             =   1665
            Width           =   2310
         End
         Begin VB.OptionButton optCommit 
            Caption         =   "Commit"
            Height          =   300
            Left            =   1080
            TabIndex        =   28
            Top             =   630
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optAbort 
            Caption         =   "Abort"
            Height          =   255
            Left            =   1110
            TabIndex        =   27
            Top             =   1020
            Width           =   1545
         End
         Begin VB.Label Label9 
            Caption         =   "Transaction action after executing UPDATE:"
            Height          =   255
            Left            =   330
            TabIndex        =   88
            Top             =   315
            Width           =   3915
         End
         Begin VB.Label Label26 
            Caption         =   $"frmDemoAppMain.frx":091E
            ForeColor       =   &H00FF0000&
            Height          =   840
            Left            =   195
            TabIndex        =   78
            Top             =   2145
            Width           =   8385
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "6. Using ExecSPbyName method"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   24
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdExecSPbyName 
            Caption         =   "Execute"
            Height          =   360
            Left            =   6675
            TabIndex        =   25
            Top             =   1680
            Width           =   2100
         End
         Begin VB.Label Label11 
            Caption         =   $"frmDemoAppMain.frx":0A73
            ForeColor       =   &H00FF0000&
            Height          =   840
            Left            =   150
            TabIndex        =   77
            Top             =   2220
            Width           =   8610
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "5. Calling Stored Procedures Through Individual Methods"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   22
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdExecCustOrderHist 
            Caption         =   "Execute"
            Height          =   330
            Left            =   6660
            TabIndex        =   23
            Top             =   1485
            Width           =   2115
         End
         Begin VB.Label Label25 
            Caption         =   $"frmDemoAppMain.frx":0B63
            ForeColor       =   &H00FF0000&
            Height          =   795
            Left            =   300
            TabIndex        =   76
            Top             =   2205
            Width           =   8430
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "4. Parameterized SQLs using Parameters Collection"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   20
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdExecSQLwPrms 
            Caption         =   "Execute"
            Height          =   360
            Left            =   6990
            TabIndex        =   21
            Top             =   1710
            Width           =   1830
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "SQL template:"
            Height          =   300
            Left            =   180
            TabIndex        =   87
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label35 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SELECT ?=ContactName FROM Customers WHERE CustomerID=? "
            Height          =   300
            Left            =   1620
            TabIndex        =   86
            Top             =   360
            Width           =   7140
         End
         Begin VB.Label Label24 
            Caption         =   $"frmDemoAppMain.frx":0C7B
            ForeColor       =   &H00FF0000&
            Height          =   795
            Left            =   105
            TabIndex        =   75
            Top             =   2205
            Width           =   8505
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "3. Parameterized SQLs with Embedded Parameters"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   15
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdExecSQLwEmbPrms 
            Caption         =   "Execute"
            Height          =   330
            Left            =   7080
            TabIndex        =   17
            Top             =   1725
            Width           =   1665
         End
         Begin VB.ComboBox cboOrderBy 
            Height          =   315
            ItemData        =   "frmDemoAppMain.frx":0D90
            Left            =   1590
            List            =   "frmDemoAppMain.frx":0D9D
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   900
            Width           =   1470
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "SQL template:"
            Height          =   300
            Left            =   135
            TabIndex        =   85
            Top             =   345
            Width           =   1230
         End
         Begin VB.Label Label33 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SELECT * FROM Orders WHERE CustomerID='%1' ORDER BY %2"
            Height          =   300
            Left            =   1575
            TabIndex        =   84
            Top             =   345
            Width           =   7140
         End
         Begin VB.Label Label23 
            Caption         =   $"frmDemoAppMain.frx":0DC2
            ForeColor       =   &H00FF0000&
            Height          =   585
            Left            =   255
            TabIndex        =   74
            Top             =   2340
            Width           =   6360
         End
         Begin VB.Label lblExecutedSQL 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1560
            TabIndex        =   63
            Top             =   1335
            Width           =   7155
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Executed SQL:"
            Height          =   300
            Left            =   225
            TabIndex        =   19
            Top             =   1395
            Width           =   1230
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Order By:"
            Height          =   300
            Left            =   555
            TabIndex        =   18
            Top             =   930
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "2. Executing SQLs, Navigating Recordset"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   -74850
         TabIndex        =   13
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdGetCustList 
            Caption         =   "Get Customers IDs"
            Height          =   330
            Left            =   6795
            TabIndex        =   14
            Top             =   1755
            Width           =   1995
         End
         Begin VB.Label Label19 
            Caption         =   $"frmDemoAppMain.frx":0E57
            ForeColor       =   &H00FF0000&
            Height          =   825
            Left            =   165
            TabIndex        =   73
            Top             =   2265
            Width           =   8190
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "1. Connecting to Database"
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   150
         TabIndex        =   7
         Top             =   1020
         Width           =   8940
         Begin VB.CommandButton cmdDsgnConnString 
            Caption         =   "..."
            Height          =   330
            Left            =   8445
            TabIndex        =   10
            ToolTipText     =   "Connection Wizard"
            Top             =   690
            Width           =   330
         End
         Begin VB.TextBox txtConnectString 
            Height          =   330
            Left            =   210
            TabIndex        =   9
            Text            =   "Provider=SQLOLEDB.1;Password="""";Persist Security Info=True;User ID=sa;Initial Catalog=Northwind"
            Top             =   690
            Width           =   8205
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "Connect"
            Height          =   330
            Left            =   6885
            TabIndex        =   8
            Top             =   1080
            Width           =   1920
         End
         Begin VB.Label Label37 
            Caption         =   "Note 1: Make sure you provide connection to MS SQL Server Northwind database, not MS Access Database!"
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   210
            TabIndex        =   89
            Top             =   2010
            Width           =   8385
         End
         Begin VB.Label Label12 
            Caption         =   $"frmDemoAppMain.frx":0F65
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   225
            TabIndex        =   12
            Top             =   2520
            Width           =   8385
         End
         Begin VB.Label Label2 
            Caption         =   "Provide connection string to existing Northwind SQL Server database, or use Wizard to construct the string"
            Height          =   285
            Left            =   180
            TabIndex        =   11
            Top             =   375
            Width           =   8130
         End
      End
   End
   Begin VB.Label Label14 
      Caption         =   "(Customer ID is used in tests. This combobox is filled by sample in page #2)"
      Height          =   300
      Left            =   3555
      TabIndex        =   6
      Top             =   105
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Customer ID:"
      Height          =   300
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   1680
   End
End
Attribute VB_Name = "frmDemoAppMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sample application demonstrating CNorthwindDB database class generated by CLASSter(tm)
'Copyright URFIN JUS (www.urfinjus.net), 2001-2002, All rights reserved.

Option Explicit
Dim mNWindDB As CNorthwindDB

'1. Connecting to database ======================================================================================
Private Sub cmdConnect_Click()
    On Error GoTo errHandler
    Dim ErrDescr As String
    Debug.Assert FalseIfWantStepIn
    Set mNWindDB = New CNorthwindDB
    Check mNWindDB.TestConnection(txtConnectString, ErrDescr), EXC_GENERAL, "Connection attempt failed: " & ErrDescr
    mNWindDB.ConnectionString = txtConnectString.Text
    mNWindDB.Connect
    Say "Connected successfully."
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdConnect_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDsgnConnString_Click()
    On Error GoTo errHandler
    Dim strTmp As String
    strTmp = txtConnectString.Text
    If DesignConnectionString(Me.hWnd, strTmp) Then txtConnectString.Text = strTmp
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdDsgnConnString_Click", , EA_NORERAISE
    HandleError
End Sub

'2. Executing SQLs, Navigating Recordset ======================================
Private Sub cmdGetCustList_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    cboCustIDs.Clear
    With NWindDB
        .ExecSQL "SELECT * FROM Customers"
        While Not .EOF
            cboCustIDs.AddItem .Value("CustomerID")
            .MoveNext
        Wend
    End With
    If cboCustIDs.ListCount > 0 Then cboCustIDs.ListIndex = 0
    MsgBox "Customers list retrieved successfully."
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdGetCustList_Click", , EA_NORERAISE
    HandleError
End Sub

'3. Parameterized SQLs with Embedded Parameters ===================================
Private Sub cmdExecSQLwEmbPrms_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    With NWindDB
        .ExecSQL "SELECT * FROM Orders WHERE CustomerID='%1' ORDER BY %2", _
            Array(SelCustID, cboOrderBy.Text)
        lblExecutedSQL.Caption = .Command.CommandText
        frmShowXML.xml = .xmlGetData
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdExecSQLwEmbPrms_Click", , EA_NORERAISE
    HandleError
End Sub

'4. Parameterized SQLs using Parameters Collection =======================
Private Sub cmdExecSQLwPrms_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    With NWindDB
        .AddSQLParam , "ContName", adParamOutput, adVarChar, 50
        .AddSQLParam SelCustID, , , adVarChar, 10
        .ExecSQL "SELECT ?=ContactName FROM Customers WHERE CustomerID=? "
        Say "Contact Name: " & .ParamValue("ContName")
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdExecSQLwPrms_Click", , EA_NORERAISE
    HandleError
End Sub

'5. Calling Stored Procedures Through Individual Methods ==================================
Private Sub cmdExecCustOrderHist_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    With NWindDB
        .ExecCustOrderHist SelCustID
        frmShowXML.xml = .xmlGetData
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdExecSalesByYear_Click", , EA_NORERAISE
    HandleError
End Sub

'6 : Using ExecSPbyName method ===================================================
Private Sub cmdExecSPbyName_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    With NWindDB
        .ExecSPbyName SP_CUSTORDERHIST, PRM_CUSTORDERHIST, SelCustID
        frmShowXML.xml = .xmlGetData
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdExecSPbyName_Click", , EA_NORERAISE
    HandleError
End Sub

'7. Using transactions ============================================================
Private Sub cmdIncCustOrders_Click()
    On Error GoTo errHandler
    With NWindDB
        .BeginTransaction
        .ExecSQL "UPDATE Orders SET Freight=Freight + 1 " & _
              " WHERE CustomerID='%1'", SelCustID
        If optCommit.Value Then .SetComplete Else .SetAbort
        .ExecSQL "Select * from Orders WHERE CustomerID='%1'", SelCustID
        frmShowXML.xml = .xmlGetData
    End With
    Exit Sub
errHandler:
    ErrorIn "frmCLASSterSampleMain.cmdIncCustOrders_Click", , EA_NORERAISE
    HandleError
End Sub

'8. Error Handling ==================================================================
    'The first example demonstrates automatic actions inside database class:
    'producing and keeping error report; re-raising errors; aborting transactions
    'We recommend doing this inside the class ONLY if you call it from VB script
    '(without re-raising errors).
    'We are using another instance of the class as we tampering with
    'class Options; we show how to use externally set Connection object.
Private Sub cmdExecWithErrors1_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    Dim tmpDB As CNorthwindDB
    Set tmpDB = New CNorthwindDB
    Set tmpDB.Connection = NWindDB.Connection
    With tmpDB
        .Options(CTRO_ERR_RAISE) = (chkErrRaise.Value = vbChecked)
        .Options(CTRO_ERR_SETABORT) = (chkErrAbort.Value = vbChecked)
        If chkErrTrans Then .BeginTransaction
        If .ExecSQL(txtErrSQL.Text) Then
            'ExecSQL returned True, meaning success
            If chkErrTrans Then .SetComplete
            Say "SQL statement executed successfully"
            Else
            'ExecSQL returned False, meaning error occured, and it was not re-raised
            Say "Error occurred; it was NOT re-raised. About to show error report."
            frmShowError.ErrorReport = .ErrorDescription
        End If
    End With
    Exit Sub
errHandler:
    Say "Error occurred; it was re-raised. About to show error report."
    ErrorIn "frmDemoAppMain.cmdExecWithErrors1_Click", , EA_NORERAISE, tmpDB
    HandleError
End Sub

'The second example demonstrates recommended way of handling errors in
'any environment other than VBScript.
Private Sub cmdExecWithErrors2_Click()
    On Error GoTo errHandler
    Dim EA As Long
    Debug.Assert FalseIfWantStepIn
    EA = EA_NORERAISE
    With NWindDB
        If chkErrTrans Then
            EA = EA Or EA_ROLLBACK 'Set flag to rollback transaction on error
            .BeginTransaction
        End If
        .ExecSQL txtErrSQL.Text
         If chkErrTrans Then .SetComplete
        Say "SQL statement executed successfully"
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdExecWithErrors2_Click", , EA, mNWindDB
    HandleError
End Sub

'9. Batch Execution, Multiple Recordsets ==========================================================
Private Sub cmdExecBatch_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    With NWindDB
        .BatchBegin
        .ExecSQL "SELECT * FROM Customers WHERE CustomerID='%1'", SelCustID
        .ExecCustOrderHist SelCustID
        .BatchExec
        txtBatch.Text = .BatchText
        frmShowXML.xml = .xmlGetDataAll
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdExecBatch_Click", , EA_NORERAISE
    HandleError
End Sub

'10. Automatic Fields Formatting ============================================================
Private Sub cmdShowOrderFmtted_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    With NWindDB
        .FormatsClear
        .FormatSet txtFmtMoney.Text, adCurrency
        .FormatSet txtFmtDateDft.Text, adDBTimeStamp
        .FormatSet txtFmtShipDate.Text, "ShippedDate"
        .ExecSQL "SELECT TOP 1 * FROM Orders WHERE CustomerID='%1'", SelCustID
        If .HasRecords Then
            lblFoundOrder.Caption = "ID=" & .ValueFmt("OrderID") & ", OrderDate=" & .ValueFmt("OrderDate") & _
               ", ShipDate=" & .ValueFmt("ShippedDate") & ", Freight=" & .ValueFmt("Freight")
            Else
            lblFoundOrder.Caption = "(Order not found)"
        End If
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdShowOrderFmtted_Click", , EA_NORERAISE
    HandleError
End Sub

'11. XML Functionality, Basic Level ==============================================================
Private Sub cmdBuildXMLBasic_Click()
    On Error GoTo errHandler
    Debug.Assert FalseIfWantStepIn
    With NWindDB
        .BatchBegin
        .ExecSQL "SELECT * FROM Customers WHERE CustomerID='%1'", SelCustID
        .ExecCustOrderHist SelCustID
        .BatchExec
        frmShowXML.xml = .xmlGetDataAll(cboOpenTag.Text, txtRSElems.Text, txtRecElems.Text, txtCloseTag.Text)
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdBuildXMLBasic_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboOpenTag_Click()
    On Error GoTo errHandler
    If cboOpenTag.ListIndex = 0 Then
        txtCloseTag.Text = ""
        Else
        txtCloseTag.Text = "</CustomerInfo>"
    End If
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cboOpenTag_Click", , EA_NORERAISE
    HandleError
End Sub

'12. XML Functionality, Advanced Level
Private Sub cmdBuildXMLBasic2_Click()
    On Error GoTo errHandler
    With NWindDB
       .ExecSQL "Select *, GETDATE() As CurrDate from Customers Where country='%1'", "Germany"
        .FormatSet "DD-MMM-YYYY", adDBTimeStamp
       frmShowXML.xml = .xmlGetDataAdv("Customers", txtNodeSpecs.Text, CLng(txtSkipRecs.Text), CLng(txtIncRecs.Text))
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdBuildXMLBasic2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboNodeSpecs_Click()
    On Error GoTo errHandler
    txtNodeSpecs.Text = cboNodeSpecs.Text
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cboNodeSpecs_Click", , EA_NORERAISE
    HandleError
End Sub


'Miscellaneous methods ============================================================================
Public Property Get NWindDB() As CNorthwindDB
    On Error GoTo errHandler
    Check Not mNWindDB Is Nothing, EXC_GENERAL, "Not connected to database."
    Check mNWindDB.Connected, EXC_GENERAL, "Not connected to database."
    Set NWindDB = mNWindDB
    Exit Property
errHandler:
    ErrorIn "frmDemoAppMain.NWindDB"
End Property

Private Function SelCustID() As String
    On Error GoTo errHandler
    Check cboCustIDs.Text <> "", EXC_GENERAL, "Customers IDs are not loaded. Please execute sample on page 2."
    SelCustID = cboCustIDs.Text
    Exit Function
errHandler:
    ErrorIn "frmDemoAppMain.SelCustID"
End Function

Private Sub HandleError()
    If InException Then
        Say Err.Description, "Exception"
    Else
        frmShowError.ErrorReport = ErrReport
    End If
End Sub

Private Function FalseIfWantStepIn() As Boolean
    FalseIfWantStepIn = Not (chkStopInProc.Value = vbChecked)
End Function

Private Sub cmdNextPage_Click()
    On Error GoTo errHandler
    With sstMain
        .Tab = IIf(.Tab = .Tabs - 1, 0, .Tab + 1)
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdNextPage_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrevPage_Click()
    On Error GoTo errHandler
    With sstMain
        .Tab = IIf(.Tab = 0, .Tabs - 1, .Tab - 1)
    End With
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.cmdPrevPage_Click", , EA_NORERAISE
    HandleError
End Sub

Private Function DesignConnectionString(ByVal HWin As Long, ByRef AConnectstring As String) As Boolean
    On Error GoTo errHandler
    Dim Conn As ADODB.Connection, DLink As DataLinks
    Set DLink = New DataLinks
    DLink.hWnd = HWin
    If Trim(AConnectstring) = "" Then
        Set Conn = DLink.PromptNew
        If Not Conn Is Nothing Then
            AConnectstring = Conn.ConnectionString
            DesignConnectionString = True
        End If
        Else
        Set Conn = New Connection
        Conn.ConnectionString = AConnectstring
        If DLink.PromptEdit(Conn) Then
            AConnectstring = Conn.ConnectionString
            DesignConnectionString = True
        End If
    End If
    Exit Function
errHandler:
    ErrorIn "frmDemoAppMain.DesignConnectionString(HWin,AConnectstring)", Array(HWin, AConnectstring)
End Function

Public Sub Say(ByVal Msg As String, Optional ByVal Caption = "CLASSter Demo")
    MsgBox Msg, , Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload frmShowError
    Unload frmShowXML
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    cboOrderBy.ListIndex = 0
    cboNodeSpecs.ListIndex = 0
    sstMain.Tab = 0
    Exit Sub
errHandler:
    ErrorIn "frmDemoAppMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub


