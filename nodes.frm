VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Node Example"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Style"
      Height          =   1720
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   4695
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "nodes.frx":0000
         Left            =   1200
         List            =   "nodes.frx":000A
         TabIndex        =   10
         Text            =   "0 - tvwAutomatic"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "nodes.frx":002F
         Left            =   1200
         List            =   "nodes.frx":0039
         TabIndex        =   8
         Text            =   "0 - tvwTreeLines"
         Top             =   960
         Width           =   3375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "nodes.frx":0061
         Left            =   1200
         List            =   "nodes.frx":006B
         TabIndex        =   4
         Text            =   "0 - ccNone"
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "nodes.frx":008E
         Left            =   1200
         List            =   "nodes.frx":00AA
         TabIndex        =   3
         Text            =   "6 - tvwTreelinesPlusMinusText"
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Label Edit"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Line Style"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Style"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Boarder Style"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Treeview"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5318
         _Version        =   393217
         Style           =   6
         HotTracking     =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.Menu nodes 
      Caption         =   "nodes"
      Visible         =   0   'False
      Begin VB.Menu bold 
         Caption         =   "Bold"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''
'node example by izek shamilov'
'for www.planetsourcecode.com '
'''''''''''''''''''''''''''''''

Option Explicit ' force variable declaration
Private Sub bold_Click()
'declare the 2 variables that will be used
Dim i%, j%

'assign 0 to the intiger that holds total number of
'bold nodes
j% = 0

'a simle for loop from 1 to node count in treeview1
'same as for i = 0 to list1.listcount - 1 if using a
'ListBox
For i = 1 To TreeView1.nodes.Count

    'check if node is bold and if it is increase counter by 1
    If TreeView1.nodes.Item(i).bold = True Then j% = j% + 1
Next i

'display a message box showing how many nodes out of
'total number of nodes are bold
Call MsgBox(j% & " out of " & TreeView1.nodes.Count & " nodes are bold", vbOKOnly, "Bold Nodes")
End Sub
Private Sub Combo1_Click()
'used to change the style of treeview1
'i made default style 6
TreeView1.Style = Combo1.ListIndex
End Sub
Private Sub Combo2_Click()
'used to change the boarder style of treeview1
TreeView1.BorderStyle = Combo2.ListIndex
End Sub
Private Sub Combo3_Click()
'used to change the linestyle of treeview1
TreeView1.LineStyle = Combo3.ListIndex
End Sub
Private Sub Combo4_Click()
'used to change the labeledit of treeview1
TreeView1.LabelEdit = Combo4.ListIndex
End Sub
Private Sub Form_Load()
'declare nodx as a node
Dim nodX As Node

'add the Main node to the tree view
Set nodX = TreeView1.nodes.Add(, , "M", "Main")

'make sure that this node is visible and expanded
nodX.EnsureVisible

Set nodX = TreeView1.nodes.Add("M", tvwChild, "R", "Root")
'add the root node and make it child of the main node
'tvwChild relationship makes a node child of another node
'other relationships are
'tvwFirst - First Sibling
'tvwLast - Last Sibiling
'tvwNext - Next Sibiling
'tvwPrevious - Previous Sibiling

'add 4 child nodes to the root node each 1 with its
'individual key and ensure that they are all visible
Set nodX = TreeView1.nodes.Add("R", tvwChild, "C1", "Child 1")
Set nodX = TreeView1.nodes.Add("R", tvwChild, "C2", "Child 2")
Set nodX = TreeView1.nodes.Add("R", tvwChild, "C3", "Child 3")
Set nodX = TreeView1.nodes.Add("R", tvwChild, "C4", "Child 4")
nodX.EnsureVisible
   
   
Set nodX = TreeView1.nodes.Add("M", tvwChild, "F", "Folder")
'add another child node to the Main node, same as Root
'and then add 4 child nodes to it. Same as above
Set nodX = TreeView1.nodes.Add("F", tvwChild, "F1", "Sub Folder 1")
Set nodX = TreeView1.nodes.Add("F", tvwChild, "F2", "Sub Folder 2")
Set nodX = TreeView1.nodes.Add("F", tvwChild, "F3", "Sub Folder 3")
Set nodX = TreeView1.nodes.Add("F", tvwChild, "F4", "Sub Folder 4")
nodX.EnsureVisible

End Sub
Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'this is an example of a simple popup menu
'you make a menu and you make "parent" menu invisble
'in this example "parent" menu is called nodes
'so you can either change its visible property in the menu
'editor or in form1 load put nodes.visible = false
'that ensures that your menu and submenus of that menu are
'not visible at runtime
'this event is trigged when you click on the treeview1
'you check if it was clicked on with right or left button
'and if its right button then you call the popupmenu for
'our parent menu to display all of its sub menus

'1 - left button
'2 - right button
'3 - middle button?

'if right button then call the nodes menu
If Button = 2 Then Call PopupMenu(nodes)
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'this event has 2 parts

'part1
'when a node is clicked we set the caption of the form
'to whatever it was plus key name and node name of the
'selected node
Form1.Caption = "Node Example - " & Node.Key & " - " & Node

'part2
'we check if the clicked node is bold is true
'if it is then we set it to false, if its false we set it
'to true. Note. Exit Sub at the end of the if statement is
'required for it to work or else you will not be able
'to make bold = false because every time you would
'change it to false the next line will always change it
'back to true
If Node.bold = True Then Node.bold = False: Exit Sub
If Node.bold = False Then Node.bold = True: Exit Sub
End Sub
