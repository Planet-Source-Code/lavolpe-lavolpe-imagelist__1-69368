VERSION 5.00
Begin VB.UserControl UserControl1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   ScaleHeight     =   660
   ScaleWidth      =   1665
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' An example of using a single-purpose usercontrol to also act as
' an custom Image List control too.  This allows you to guarantee
' you have a method of storing images that does not rely on the
' common controls imagelist control, prevents you from having to
' distribute a separate control, allows you full flexibility to
' tweak the imagelist control to your special requirements.

' Combining your specific purpose control and the custom imagelist
' control does require a couple of specifc routines listed below,
' along with some special code that will reside in normal routines
' :: custom routines are
'       Mode Get/Let property
'       ImageLists Get/Set property
'       ImageList & ImageID Get/Let properties
'       ImageListFromControl routine
'       eMode enumeration adjusted to your needs
'       The ShowWindow API declaration
'       The m_ImageLists, m_Mode variables, m_ImageListName, m_ImageID variables
'       a Refresh routine which all drawn controls should probably have
'       -- of course, the ppgImageList property page and the various classes and parsing module
' :: existing routines that will need imagelist related code are
'       Usercontrol.ReadProperties & WriteProperties event (sample given)
'       Usercontrol.Show event to hide during runtime & draw if AutoRedraw=True (sample given)
'       Usercontrol.Paint event to draw if AutoRedraw=False (sample given)
' :: required code change
'       Entire Project: Search & Replace:  UserControl1 to YourUCname < change appropriately
' :: referencing the imagelist control by other controls
'    there are a couple of ways to do it. Regardless of which way you
'    choose, do not perform the task if this conttol's mode is ImageList
'    1. You can add an option to your control's property page to select
'       the usercontrol that contains the imagelist. Recommend enumerating
'       the controls as above and placing them in a dropdown box for your
'       user to select from. Here you may have one or several imagelist controls
'    2. You can add a usercontrol property so user can type the imagelist control
'       name and then create a reference from it (i.e., Controls(imgLstName) )
'     Now that you have a reference to the control, you can call the control's ImageLists
'     property. See FindImageList example towards end of module. The cImageLists
'     class has a render function where you pass the target DC, the imagelist item
'     index or key, and other parameters for rendering. The class is the only thing
'     you will want to keep a reference to, not the control acting as an image list

Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private m_Mode As eMode '-1=imagelist
Private m_ImageListName As String       ' the name of the imagelist control will use
Private m_ImageID As String             ' The numerical ID or Key of the image to be displayed
Private m_ImageLists As cImageLists
Public Enum eMode
    lvb_ImageList = -1          ' required
    lvb_CommandButton = 0       ' whatever your usercontrol is/does
    lvb_CheckBox = 1
    lvb_OptionButton = 2
End Enum

' when sharing your usercontrol with the imagelist, the property page needs
' to get the uc's imagelist class.  Friend properties for that purpose
' Also, this property can be used for your non-imagelist uc to find the imagelist.
' Friend vs Public so it is not exposed to the user
Friend Property Get ImageLists() As cImageLists
    Set ImageLists = m_ImageLists
End Property
Friend Property Set ImageLists(newIL As cImageLists)
    PropertyChanged "Mode"  ' only used so the property page can force this uc to update changes
End Property

Public Property Let Mode(newMode As eMode)
    ' to share imagelist control and your control as same control....
    ' 1. Only allow property to be changed during design time
    ' 2. Depending on setting, change property pages and/or redraw, reset, etc
    
    If Ambient.UserMode = True Then Exit Property ' design time property only
    
    ' validate within range of Mode values
    If newMode < lvb_ImageList Or newMode > lvb_OptionButton Then Exit Property
    
    If Not newMode = m_Mode Then            ' only do if property changed
        If m_Mode = lvb_ImageList Then
            ' was imagelist before, warn user that changing mode clears images
            If MsgBox("Changing the mode will delete all images currently mananged by the imagelist mode. Continue?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then Exit Property
            ' clean up, clear the class that contains the DIB and imagelist settings
            Set m_ImageLists = Nothing
            ' now add propertypages that apply to your uc when not an imagelist
            ' if you don't have any property pages, set to vbNullString instead
            UserControl.PropertyPages(0) = vbNullString
        End If
        
        If newMode = lvb_ImageList Then
            ' ok, the only requirement is that PNGs must be able to be created
            ' Validate & warn if not. When not in design mode, this is not applicable
            ' Reason? PNGs are used to save the imagelists. Without that, size of
            ' imagelists in your project could be absolutely huge
            Dim c As c32bppDIB
            Set c = New c32bppDIB
            If Not c.isGDIplusEnabled Then
                If Not c.isZlibEnabled Then
                    MsgBox "The image lists cannot be stored without the appropriate DLLs installed." & vbCrLf & _
                    "1. GDI+ can be downloded from Microsoft, or" & vbCrLf & _
                    "2. zLIB can be downloaded from www.zlib.net." & vbCrLf & _
                    "The DLL should be located in your system path", vbExclamation + vbOKOnly, "Developer Warning"
                    Exit Property
                End If
            End If
            Set c = Nothing
            m_Mode = newMode                    ' set the new mode
            ' if converting to an imagelist, then set the property page & welcome user
            Set m_ImageLists = New cImageLists              ' create new imagelist class
            UserControl.PropertyPages(0) = "ppgImageList"   ' set property page
            ' remove any previous property pages
            Do Until UserControl.PropertyPages(1) = vbNullString
                PropertyPages(1) = vbNullString
            Loop
            MsgBox "Right click on the control, select properties to begin adding images", vbInformation + vbOKOnly, "Mode Changed"
        
        Else
            m_Mode = newMode                    ' set the new mode
            
            ' redraw your non-imagelist usercontrol, refresh it, whatever
            Call Refresh
            
        End If
        
        PropertyChanged "Mode"
    End If
End Property
Public Property Get Mode() As eMode
    Mode = m_Mode
End Property

Public Property Let ImageList(imgList As String)
    ' allow user to set the imagelist to be used
    ' Indexed controls must be passed like so:  controlName(index)
    If Not m_Mode = lvb_ImageList Then
        m_ImageListName = imgList
        Call ImageListFromControl(True)     ' validate the control & get imagelist ref
        PropertyChanged "ImageList"
        ' redraw your control
        Call Refresh
    End If
End Property
Public Property Get ImageList() As String
    If m_Mode = lvb_ImageList Then
        ImageList = "ImageList Source"
    Else
        ImageList = m_ImageListName
    End If
End Property

Public Property Let ImageID(iID As String)
    ' allow user to specify which image ID from the list
    ' iID can be a numerical index or a Key
    If Not m_Mode = lvb_ImageList Then
        m_ImageID = iID
        PropertyChanged "ImageID"
        ' redraw your control
        Call Refresh
    End If
End Property
Public Property Get ImageID() As String
    If m_Mode = lvb_ImageList Then
        ImageID = m_ImageLists.Count & " Images"
    Else
        ImageID = m_ImageID
    End If
End Property


Public Sub Refresh()
    ' simple refresh routine. Obviously yours will be far more detailed & personal
    Cls
    If Not m_ImageLists Is Nothing Then ' don't draw the graphic
        Dim X As Long, Y As Long
        X = ScaleWidth \ 2
        Y = ScaleHeight \ 2
        ' passing an invalid m_ImageID will cause drawing to fail here, but will not error
        ' You can test the return value which will be True, or False if failed
        m_ImageLists.Render m_ImageID, UserControl.hDC, X, Y, , , , , , True
        ' ^^ note that there should some advanced backbuffer drawing taking place.
        '    The above is just a simple example
    End If
End Sub


Private Function ImageListFromControl(bShowError As Boolean) As Boolean

    On Error Resume Next
    ' pass bShowError if you want a messagebox error displayed.
    ' Recommending setting for True in Property Let ImageList & False in UserControl_Show
    
    If m_Mode = lvb_ImageList Then
        ImageListFromControl = True
        Exit Function
    ElseIf m_ImageListName = vbNullString Then
        Set m_ImageLists = Nothing
        ImageListFromControl = True
        Exit Function
    End If
    
    Dim c As UserControl1   ' < change name to your usercontrol name
    Dim cName As String, cIndex As Long
    Dim iRaiseError As Long
    
    ' parse the imagelist name
    If Right$(m_ImageListName, 1) = ")" Then
        cName = Left$(m_ImageListName, InStr(m_ImageListName, "(") - 1)
        cIndex = Val(Mid$(m_ImageListName, InStr(m_ImageListName, "(") + 1))
        Set c = Parent.Controls(cName)(cIndex)
    Else
        cName = m_ImageListName
        Set c = Parent.Controls(cName)
    End If
    If Err Then                 ' was passed name valid?
        iRaiseError = 1         ' nope or it wasn't a UserControl1 object
    Else
        cIndex = c.Mode         ' got right object, does it have right methods?
        If Err Then
            iRaiseError = 2     ' nope, doesn't have the Mode method
        Else
            Set m_ImageLists = c.ImageLists
            If Err Then iRaiseError = 2 ' nope, doesn't have the ImageLists method
        End If                          ' or method does not return a cImageLists class
    End If
    
    On Error GoTo 0
    If iRaiseError = 0 Then
        ImageListFromControl = True
    Else
        Err.Clear
        Set m_ImageLists = Nothing
        m_ImageListName = vbNullString
        If bShowError Then
            Select Case iRaiseError
            Case 1
                Err.Raise vbObject Or 53, Ambient.DisplayName, "Invalid Image List Name"
            Case 2
                Err.Raise vbObject Or 53, Ambient.DisplayName, "Not an Image List"
            End Select
        End If
    End If

End Function

Private Sub UserControl_Initialize()
    ScaleMode = vbPixels                ' recommendation when owner drawn anything
End Sub

Private Sub UserControl_Paint()
    ' this event only gets fired if AutoRedraw=False.
    ' Otherwise this event can be removed completely
    If m_Mode = lvb_ImageList Then
        UserControl.CurrentY = 0        ' reset
        UserControl.CurrentX = 0
        UserControl.Print "Custom" & vbCrLf & "ImageList"
    Else
        ' custom draw your uc
        Call Refresh
    End If
End Sub

Private Sub UserControl_Show()
    If m_Mode = lvb_ImageList Then
        ' when running, we don't want the imagelist uc displayed
        If Ambient.UserMode = True Then
            ShowWindow UserControl.hwnd, 0&     ' hide during runtime
        Else
            ' change to suit taste. This is how you want it to appear in design mode
            ' make it as pretty or as plain as you wish
            On Error Resume Next
            Appearance = 0
            BorderStyle = 1
            BackStyle = 1
            BackColor = vbInfoBackground
            ForeColor = vbInfoText
            ' if you change the caption below, change it in the Usercontrol_Paint event too
            If AutoRedraw = True Then UserControl.Print "Custom" & vbCrLf & "ImageList"
        End If
        
    Else
            Call ImageListFromControl(False)    ' must call to set the image list reference
            
            ' add code needed for your usercontrol
            ' draw it first time if AutoRedraw=True. Example:
            Call Refresh
            
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        .WriteProperty "Mode", m_Mode, lvb_CommandButton
        If m_Mode = lvb_ImageList Then
            
            ' the uc is in imagelist mode....
            Dim dat() As Byte, imgCount As Long
            Dim Index As Long, ItemNr As Long
            
            ' get the imagelist from the class as a single PNG file
            ' each imagelist (one per same-sized images) will be separate PNGs.
            Index = 1
            Do Until m_ImageLists.ExportImageList(Index, dat(), imgCount) = False
                ItemNr = ItemNr + 1
                .WriteProperty "ilDat" & ItemNr, dat()
                .WriteProperty "ilHdr" & ItemNr, imgCount
            Loop
            ' now get the individual image item properties: tag, key, etc
            ItemNr = 0: Index = 1
            Do Until m_ImageLists.ExportImageItemData(Index, dat()) = False
                ItemNr = ItemNr + 1
                .WriteProperty "ilItem" & ItemNr, dat()
            Loop
            ' now get any additional properties
            m_ImageLists.ExportMiscProperties dat
            .WriteProperty "ilMisc", dat
            Erase dat
            
        Else    ' uc is in your mode, save whatever properties you need saved
        
            .WriteProperty "ilName", m_ImageListName, vbNullString
            .WriteProperty "imgID", m_ImageID, vbNullString
            
        End If
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        
        m_Mode = .ReadProperty("Mode", lvb_CommandButton)
        If m_Mode = lvb_ImageList Then ' image list mode
            
            Dim dat() As Byte, imgCount As Long
            Dim ItemNr As Long
            
            Set m_ImageLists = New cImageLists              ' initialize a imagelisit class
            UserControl.PropertyPages(0) = "ppgImageList"   ' add only property page
            ' control is in imagelist mode. Set its property page & remove any
            ' additional default pages
            Do Until UserControl.PropertyPages(1) = vbNullString
                UserControl.PropertyPages(1) = vbNullString
            Loop
            ' no get the imagelist from the propertybag and pass it to the class
            ItemNr = 0
            Do                              ' first the images (a single PNG file per image list)
                ItemNr = ItemNr + 1
                dat = .ReadProperty("ilDat" & ItemNr, vbNullString) ' imagelist specific property
                If UBound(dat) = -1 Then Exit Do
                imgCount = .ReadProperty("ilHdr" & ItemNr, 0&)      ' imagelist specific property
                m_ImageLists.ImportImageList dat, imgCount
            Loop
            ItemNr = 0
            Do                              ' next the image key, tag properties
                ItemNr = ItemNr + 1
                dat = .ReadProperty("ilItem" & ItemNr, vbNullString) ' imagelist specific property
                If UBound(dat) = -1 Then Exit Do
                m_ImageLists.ImportImageItemData dat
            Loop                            ' last any additional properties
            dat = .ReadProperty("ilMisc", vbNullString)             ' imagelist specific property
            m_ImageLists.ImportMiscProperites dat
            
        Else    ' mode is your uc. Add property pages as needed
                ' and include other startup code
            m_ImageListName = .ReadProperty("ilName", vbNullString)
            m_ImageID = .ReadProperty("imgID", vbNullString)
        
        End If
    End With
End Sub

Private Sub FindImageList()

' simple sample routine to locate a imagelist versions of your control
' This example would be in a propertypage where you wanted to display
' a listing of imagelist controls for the user to select from


    If Not m_Mode = lvb_ImageList Then
    
        Dim X As Long
        Dim myType As String
        
        ' here you would declare the type of control you will be looking for
        ' If it is your uc, then change UserControl1 appropriately, if it
        ' is another UC type, then change it to that type. When changing,
        ' change it in both places below.
        
        Dim tmpUC As UserControl1 ' needed to reference Friend properties
        myType = "UserControl1"
        
        On Error Resume Next
        For X = 0 To ParentControls.Count - 1
            If TypeName(ParentControls(X)) = myType Then
                If ParentControls(X).Mode = lvb_ImageList Then  ' .Mode is public, no errors
                    Set tmpUC = ParentControls(X)               ' .ImageLists is Friend, need hard reference
                    Set m_ImageLists = tmpUC.ImageLists
                    Debug.Print "imagelist found for " & Ambient.DisplayName
                    Exit For
                End If
            End If
        Next
    
    End If

End Sub
