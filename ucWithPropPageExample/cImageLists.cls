VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A single class solution to create and manage one or more separate image lists.
' The ONLY requirement to use this code is that you MUST HAVE one of the following
' DLLs installed on your system:
'    1. GDI+ comes with XP and Vista, can be installed on WinNT4+ & Win98+. Download from Microsoft
'    2. zLIB can be found at www.zlib.net, is free, and is compatible with all operating systems
' Without one of those, this class' ExportImageList will fail. You won't be able to save imagelists

' A quick summary.
' :: an image list is defined as a single image that contains multiple individual images
'       of the same height and width
' :: this class can contain several image lists
' :: this class is designed to be called from both a custom usercontrol to render images
'       to and also from the supplied propertypage object: ppgImageList
' :: The various Export & Import routines are designed to serialize and deserialize the
'       imagelists and individual image attributes. See the included usercontrol's
'       ReadProperties and WriteProperties to see how this is done
' :: You will be able to see every routine within this class from any object that
'       references this class.  Calling those routines can result in lost data
'       if you do not understand what they are doing.

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Type imageListStruct    ' Each size image has its own image list which can contain many separate images
    Count As Long               ' number of same size images on this imagelist
    Width As Long               ' width of individual image; cached cause used often
    Host As c32bppDIB           ' DIB holding the imagelist pixels
End Type
Private Type imageItemStruct    ' collection item details
    Slot As Long                ' position in imagelist strip where image falls
    ListIndex  As Long          ' array index to which imagelist this item falls
    Key As String               ' reverse lookup for collection items, if used
    Tag As String               ' user supplied data
End Type

Private m_Lists() As imageListStruct    ' collection of imagelist DIBs
Private m_Items() As imageItemStruct    ' collection of imagelist item details
Private cLists As Collection            ' key-to-index crossreference for imagelists
Private cItems As Collection            ' key-to-index crossrefernce for image items
Private m_ItemCount As Long             ' number of total image items
' Note. The cItems collection has entries only when a key is used. There
' may be more m_Items entries than cItem entries.


' used only by image list property page to allow undoing
Private c_ItemsCloned() As imageItemStruct
Private c_ListsCloned() As imageListStruct
Private cListsCloned As Collection
Private cItemsCloned As Collection
Private c_ItemCount As Long

' properties used by the property page that are desired to be persisted
Private ppgProps() As Byte  ' see ppgSetProperties

Public Function Render(ByVal Key As Variant, _
                        ByVal hDC As Long, _
                        ByVal X As Long, _
                        ByVal Y As Long, _
                        Optional ByVal Width As Long = 0&, _
                        Optional ByVal Height As Long = 0&, _
                        Optional ByVal Lightness As Long = 0&, _
                        Optional ByVal Opacity As Long = 100&, _
                        Optional ByVal grayScale As Long = eGrayScaleFormulas.gsclNone, _
                        Optional ByVal FromCenterPt As Boolean = False) As Boolean

    ' common use: renders any imagelist item to passed DC using passed parameter options
    ' This is the only routine that usercontrols should be calling on a routine basis.
    ' Other routines that can serve some purposes are
    '   the Count property which is read only
    '   the IndexFromKey property which is read only
    '   the 3 Export & Import functions designed specifically for a control acting as an imagelist
    '   the AddImage RemoveImage & UpdateImage functions to modify imagelists at runtime
    
    ' Parameters. See c32bppDIB class' Render function for more details.
    '             This routine simply calls that class' Render, passing it the
    '             correct offsets for the requested image item
    ' Key :: a numerical index or Key value of the imagelist item to render (not case sensitive)
    ' hDC :: the target DC to draw to
    ' X,Y :: the X,Y coordinates to draw to on target DC
    ' remaining parameters are optional.
    ' Width :: Width to draw to, default is the imagelist item width
    ' Height :: Height to draw to, default is the imagelist item height
    ' Lightness :: adds/subtracts intensity of rendered image
    '              Lightness can be used to create disabled or hot images on the fly
    '              Rendering normal image at -5 then at 0 for hot image produces nice effect
    '              Rendering normal image at -10 or grayscaling can produce nice disabled effect
    ' Opacity :: percentage of opaqueness image is drawn at
    ' grayScale :: one of several formulas to render image in
    ' FromCenterPt :: if true, image is rendered, centered around passed X,Y
    '                 else image is rendered from passed X,Y
    
    If cLists Is Nothing Then Exit Function
    
    On Error Resume Next
    Dim Index As Long
    
    If IsNumeric(Key) Then
        Index = CLng(Val(Key))
    Else
        Index = cItems.Item(Key)
    End If
    If Index < 1 Or Index > m_ItemCount Then
        If Err Then Err.Clear
        Exit Function
    End If
    
    On Error GoTo 0
    With m_Lists(m_Items(Index).ListIndex)
        If Width < 1 Then Width = .Width
        If Height < 1 Then Height = .Host.Height
        Render = .Host.Render(hDC, X, Y, Width, Height, _
            (m_Items(Index).Slot - 1) * .Width, 0&, .Width, .Host.Height, Opacity, , , , grayScale, Lightness, , FromCenterPt)
    End With
End Function

Public Property Get IndexFromKey(Key As Variant) As Long
    Dim Index As Long
    
    On Error Resume Next
    If IsNumeric(Key) Then
        Index = CLng(Val(Key))
    Else
        Index = cItems.Item(Key)
    End If
    If Index < 1 Or Index > m_ItemCount Then
        If Err Then Err.Clear
        Exit Property
    Else
        IndexFromKey = Index
    End If
End Property


Public Sub SetCloneMode(Clone As Boolean, bRestore As Boolean)
    ' propertypage use only:
    ' This routine copies the this class to temporary arrays/collection objects
    ' so that if user hits Cancel, we can undo
    Dim Looper As Long
    
    If Clone Then
        ' begin cloning
        Set cItemsCloned = New Collection
        If Not cLists Is Nothing Then
            Set cListsCloned = New Collection
            c_ItemsCloned = m_Items             ' copy array
            c_ListsCloned = m_Lists             ' copy array
            c_ItemCount = m_ItemCount           ' copy count
            For Looper = 1 To cLists.Count      ' copy collection entries
                Set c_ListsCloned(Looper).Host = Nothing
                m_Lists(Looper).Host.CopyImageTo c_ListsCloned(Looper).Host
                cListsCloned.Add Looper, m_Lists(Looper).Width & "x" & m_Lists(Looper).Host.Height
            Next
            For Looper = 1 To cItems.Count      ' copy collection entries
                cItemsCloned.Add cItems.Item(Looper), m_Items(cItems.Item(Looper)).Key
            Next
        End If
    Else
        If bRestore Then            ' user wants to save changes, else Cancel clicked
            Set cItems = Nothing                ' clear old items
            Set cLists = Nothing
            Erase m_Lists
            Erase m_Items
            Set cItems = cItemsCloned           ' copy updated items
            Set cLists = cListsCloned
            m_Items = c_ItemsCloned
            m_Lists = c_ListsCloned
            m_ItemCount = c_ItemCount
        End If
        Set cItemsCloned = Nothing              ' clear cloned items
        Set cListsCloned = Nothing
        Erase c_ItemsCloned
        Erase c_ListsCloned
        c_ItemCount = 0&
    End If
End Sub

Public Property Get Count() As Long
    ' returns total number of imagelist items on all image lists
    Count = m_ItemCount
End Property

Public Function ImageProps(ByVal Index As Long, _
                    Key As String, Width As Long, Height As Long, _
                    Tag As String) As Boolean

    ' propertypage use only: returns attributes of a specific imagelist item
    If m_ItemCount = 0 Then Exit Function
    If Index < 1 Or Index > m_ItemCount Then Exit Function
    
    With m_Items(Index)
        Key = .Key
        Width = m_Lists(.ListIndex).Width
        Height = m_Lists(.ListIndex).Host.Height
        Tag = .Tag
    End With
    
End Function

Public Function ImageListSizes(outArray() As Long)
    ' propertypage use only: returns widhts/heights of all imagelists
    Dim Index As Long
    If cLists Is Nothing Then Exit Function
    If cLists.Count = 0 Then Exit Function
    
    ReDim outArray(0 To cLists.Count - 1, 0 To 1)
    For Index = 1 To cLists.Count
        outArray(Index - 1, 0) = m_Lists(Index).Width
        outArray(Index - 1, 1) = m_Lists(Index).Host.Height
    Next
    ImageListSizes = cLists.Count
End Function

Public Sub RenderRangeHorizontal(ByVal StartItem As Long, ByVal StopItem As Long, hDC As Long, ByVal X As Long, ByVal Y As Long, Width As Long, Height As Long, Optional Lightness As Long = 0&, Optional Opacity As Long = 100, Optional grayScale As Long = gsclNone)

    ' propertypage use only: renders a horizontal strip used in imagelist
    If cLists Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim Index As Long
    
    
    Dim iItem As imageItemStruct, iList As imageListStruct
    
    If StopItem > m_ItemCount Then StopItem = m_ItemCount
    
    For StartItem = StartItem To StopItem
        iItem = m_Items(StartItem)
        iList = m_Lists(iItem.ListIndex)
        With iList
            .Host.Render hDC, X, Y, Width, Height, _
                (iItem.Slot - 1) * .Width, 0&, .Width, iList.Host.Height, Opacity, , , , grayScale, Lightness
        End With
        X = X + Width
    Next


End Sub

Public Function AddImage(Host As c32bppDIB, Width As Long, Height As Long, Optional Key As Variant, Optional Tag As String, Optional ScaleToSize As Boolean) As Boolean
    
    ' common use, can be called from UC or propertypage
    ' Adds an image to the end of the imagelist. Creates new imagelist if needed
    
    ' see if we already have this key
    On Error Resume Next
    If Not IsMissing(Key) Then
        If IsNumeric(Key) Then Exit Function    ' invalid key
        If Key <> vbNullString Then cItems.Add m_ItemCount + 1, Key
        If Err Then
            Err.Clear    ' duplicated key
            Exit Function
        End If
    End If
    
    Dim iList As imageListStruct
    Dim iItem As imageItemStruct
    
    Dim listKey As String, Index As Long, X As Long, Y As Long
    Dim Looper As Long, oldImage() As Byte, newImage() As Byte
    
    
    ' locate the correct imagelist. Set key for collection lookup
    listKey = Width & "x" & Height
    If cLists Is Nothing Then               ' no image list(s) established yet
        Set cLists = New Collection
    Else
        Index = cLists(listKey)             ' do we have it?
        If Err Then Err.Clear               ' clear error if we didn't have it
    End If
        
    On Error GoTo 0                         ' stop ignoring errors
    If Index = 0 Then                       ' this image list doesn't exist yet, create it
        
        With iList
            Set .Host = New c32bppDIB       ' create new imagelist DIB & initialize it
            .Host.InitializeDIB Width, Height
            .Count = 1                      ' number images on this list
            .Width = Width                  ' and individual image item dimensions
            If (Host.Width = Width And Host.Height = Height) Then
                ' if no resizing is needed, we can simply copy it directly to our imagelist now
                CopyMemory ByVal .Host.BitsPointer, ByVal Host.BitsPointer, Height * Host.scanWidth
            End If
        End With
        Index = cLists.Count + 1            ' array index of the imagelist
        ReDim Preserve m_Lists(1 To Index)  ' add imagelist to our array
        cLists.Add Index, listKey           ' add key lookup
    
    Else
    
        iList = m_Lists(Index)              ' we have the imagelist & need to add to it
        With iList
            .Count = .Count + 1                             ' increment number of images in this list
            Call .Host.GetDIBbits(oldImage())               ' get original bytes
            .Host.InitializeDIB .Width * .Count, .Host.Height    ' create new DIB of correct size
            
            ' size array and transfer old DIB scanlines to this new one. Added image will be at far right
            ReDim newImage(0 To .Host.scanWidth - 1, 0 To .Host.Height - 1)
            For Looper = 0 To .Host.Height - 1
                CopyMemory newImage(0, Looper), oldImage(0, Looper), UBound(oldImage, 1) + 1&
            Next
            If (Host.Width = Width And Host.Height = Height) Then
                ' if no resizing is needed, we can simply copy it directly to our imagelist
                Call Host.GetDIBbits(oldImage())           ' get new image
                X = Width * 4& * (.Count - 1)              ' offset in imagelist where image will appear
                For Looper = 0 To Host.Height - 1          ' copy it
                    CopyMemory newImage(X, Looper), oldImage(0, Looper), Host.scanWidth
                Next
            End If
            ' now transfer the old image bytes to the new DIB section
            CopyMemory ByVal .Host.BitsPointer, newImage(0, 0), .Host.scanWidth * .Host.Height
        End With
    End If
    m_Lists(Index) = iList                  ' update our array
    
    m_ItemCount = m_ItemCount + 1
    With iItem                              ' now update the item data
        .Slot = iList.Count                 ' which slot in image list this image resides
        .Tag = Tag                          ' user-defined
        .ListIndex = Index                  ' key to which image list item belongs to
        .Key = Key                          ' key to provide reverse lookup in collection object
    End With                                ' note: key was added at very top of routine
    ReDim Preserve m_Items(1 To m_ItemCount)
    m_Items(m_ItemCount) = iItem           ' add item to our array
    
    If Not (Host.Width = Width And Host.Height = Height) Then
        ' here we render the new image to the last slot in the imagelist
        With iList
            If ScaleToSize Then
                Host.ScaleImage .Width, .Host.Height, X, Y, eScaleOptions.ScaleToSize
            Else
                X = .Width
                Y = .Host.Height
            End If
            Host.Render 0, .Width * (.Count - 1) + (.Width - X) \ 2, (.Host.Height - Y) \ 2, X, Y, , , , , , , , .Host
        End With
    End If
    AddImage = True
    
End Function

Public Function RemoveImage(ByVal Key As Variant) As Boolean

    ' common use, can be called from UC or propertypage
    ' Removes an image to the end of the imagelist. Deletes imagelist if needed

    Dim iList As imageListStruct, iListBlank As imageListStruct
    Dim iItem As imageItemStruct, iItemBlank As imageItemStruct
    Dim Index As Long, ListIndex As Long, sKey As String
    Dim X As Long, Y As Long, slideSlot As Long
    Dim oldImage() As Byte, newImage() As Byte
    
    On Error Resume Next            ' get index of the item to be removed
    If IsNumeric(Key) Then
        Index = CLng(Val(Key))
    Else
        Index = cItems(Key)         ' error occus if not a keyed item
    End If                          ' validate index
    If Index < 1 Or Index > m_ItemCount Then
        If Err Then Err.Clear
        Exit Function
    End If
    
    slideSlot = m_Items(Index).Slot         ' position in imagelist where image falls
    ListIndex = m_Items(Index).ListIndex    ' the imagelist that contains this item
    
    If m_ItemCount = 1 Then             ' we only have one item, so no more imagelists
        Set cItems = New Collection
        Set cLists = Nothing
        Erase m_Items
        Erase m_Lists
        m_ItemCount = 0&
        RemoveImage = True
        Exit Function
        
    Else
        ' Ok, here we are going to shift all image items down, over the deleted item
        
        ' remove the key now, if applicable
        If Not m_Items(Index).Key = vbNullString Then cItems.Remove m_Items(Index).Key
        
        ' gotta be careful when shifting UDTs that contain pointers
        ' The UDT we are shifting has string pointers (Key & Tag)
        CopyMemory ByVal VarPtr(iItem), ByVal VarPtr(m_Items(Index)), Len(iItem)
        CopyMemory ByVal VarPtr(m_Items(Index)), ByVal VarPtr(m_Items(Index + 1)), Len(iItem) * (m_ItemCount - Index)
        CopyMemory ByVal VarPtr(m_Items(m_ItemCount)), ByVal VarPtr(iItemBlank), Len(iItem)
        m_ItemCount = m_ItemCount - 1               ' update count & resize array
        ReDim Preserve m_Items(1 To m_ItemCount)
        
        ' ok, now rebuild the collection item key-to-index crossreference
        ' The colletion references may be off one if we don't check
        For X = Index To m_ItemCount
            If Not m_Items(X).Key = vbNullString Then
                cItems.Remove m_Items(X).Key        ' delete collection item
                cItems.Add X, m_Items(X).Key        ' add it back with the correct index (X)
            End If
        Next
    End If
    
    If m_Lists(ListIndex).Count = 1 Then
        ' the imagelist has no more images, need to remove the imagelist
        If ListIndex = cLists.Count Then
            ' if imagelist is last in the array, easy to remove
            cLists.Remove m_Lists(ListIndex).Width & "x" & m_Lists(ListIndex).Host.Height
            ReDim Preserve m_Lists(1 To ListIndex - 1)
        Else
            ' otherwise we will swap the last list with the one being removed
            ' Maintaining order here is not important because every imagelist is keyed
            CopyMemory ByVal VarPtr(iList), ByVal VarPtr(m_Lists(ListIndex)), Len(iList)
            CopyMemory ByVal VarPtr(m_Lists(ListIndex)), ByVal VarPtr(m_Lists(cLists.Count)), Len(iList)
            CopyMemory ByVal VarPtr(m_Lists(cLists.Count)), ByVal VarPtr(iListBlank), Len(iList)
            
            ' rebuild the collection keys
            Set cLists = New Collection
            For X = 1 To UBound(m_Lists) - 1
                cLists.Add X, m_Lists(X).Width & "x" & m_Lists(X).Host.Height
            Next
            
            ' now to re-reference the items in the swapped imagelist
            ' The last imagelist was moved, therefore, its items need to be re-referenced
            Y = cLists.Count + 1
            For X = 1 To m_ItemCount
                If m_Items(X).ListIndex = Y Then
                    m_Items(X).ListIndex = ListIndex
                End If
            Next
            ReDim Preserve m_Lists(1 To cLists.Count)   ' resize our imagelist array
            
        End If
    Else
        ' this looks complicated but its not. We need to shift images over the removed image
        ' while resizing the DIB.
        With m_Lists(ListIndex)
            .Count = .Count - 1                 ' number images on this imagelist
            .Host.GetDIBbits oldImage()         ' get current imagelist bytes & build new one
            ReDim newImage(0 To .Width * 4& * .Count - 1, 0 To .Host.Height - 1)
            For Y = 0 To .Host.Height - 1
                Select Case slideSlot           ' shifting images to the left
                Case 1  ' image deleted was first, copy everything after that image
                    CopyMemory newImage(0, Y), oldImage(.Width * 4, Y), .Count * .Width * 4&
                Case .Count + 1 ' image deleted was last, copy everything before that image
                    CopyMemory newImage(0, Y), oldImage(0, Y), .Count * .Width * 4&
                Case Else   ' image in middle somewhere, copy everything before & then after
                    CopyMemory newImage(0, Y), oldImage(0, Y), (slideSlot - 1) * .Width * 4&
                    CopyMemory newImage((slideSlot - 1) * .Width * 4, Y), oldImage(slideSlot * .Width * 4&, Y), (.Count - slideSlot + 1) * .Width * 4&
                End Select
            Next
            ' resize our new imagelist & set its alpha property
            .Host.InitializeDIB .Width * .Count, .Host.Height
            .Host.Alpha = True
            ' now we copy the new imagelist bytes to the DIB
            CopyMemory ByVal .Host.BitsPointer, newImage(0, 0), .Host.scanWidth * .Host.Height
        End With
        Erase newImage  ' these no longer needed
        Erase oldImage
        
        ' now we need to adjust any image slot references in that image list
        ' All images on that same imagelist that were after the deleted image
        ' need to have their slots reduced by one
        For X = Index To m_ItemCount
            If m_Items(X).ListIndex = ListIndex Then
                If m_Items(X).Slot > slideSlot Then m_Items(X).Slot = m_Items(X).Slot - 1
            End If
        Next
    End If
    RemoveImage = True

End Function

Public Function UpdateImage(OldKey As Variant, newKey As Variant, Tag As String) As Boolean

    ' routine simply updates the Key and/or Tag property of an image
    ' Trying to update a Key, resulting in a duplicated key will return False
    ' Trying to set a numeric key will return false
    Dim Index As Long

    On Error Resume Next
    
    If IsNumeric(OldKey) Then
        Index = CLng(Val(OldKey))
    Else
        Index = cItems(OldKey)
    End If
    If Index < 1 Or Index > m_ItemCount Then
        If Err Then Err.Clear
        Exit Function
    ElseIf IsNumeric(newKey) Then
        Exit Function
    End If
    
    If Not newKey = m_Items(Index).Key Then ' key is changing
        If Len(newKey) Then
            cItems.Add Index, newKey
            If Err Then                 ' duplicated a key
                Err.Clear
                Exit Function
            End If
            ' remove oldkey if needed
            If Not m_Items(Index).Key = vbNullString Then cItems.Remove m_Items(Index).Key
            m_Items(Index).Key = newKey ' save new key
        Else                            ' removing the key
            If Not m_Items(Index).Key = vbNullString Then cItems.Remove m_Items(Index).Key
            m_Items(Index).Key = vbNullString
        End If
    End If
    m_Items(Index).Tag = Tag
    UpdateImage = True
    
End Function

Friend Function SetPropPageProps(inProps() As Byte)
    ' property page use only. Do not call from your usercontrols
    ppgProps = inProps
End Function
Friend Function GetPropPageProps(inProps() As Byte)
    inProps = ppgProps
End Function

Public Function ExportImageList(Index As Long, imgArray() As Byte, imgCount As Long) As Boolean

    ' UC use only. Called during its WriteProperties event
    ' Routine serializes each imagelist as a PNG and also returns number of images contained
    ' ImportImageList is the complimentary function

    ' Index will be incremented to next list item
    ' imgArray :: a PNG file
    ' imgCount :: number of images in this imagelist

    ' Index begins at one and continues until function returns false
    If cLists Is Nothing Then Exit Function                 ' no image lists
    If Index < 1 Or Index > cLists.Count Then Exit Function ' invalid Index passed
    
    With m_Lists(Index)
        ' use PNG format to compress the 32bpp DIB. Huge size savings
        If .Host.SaveToStream_PNG(imgArray()) Then
            imgCount = m_Lists(Index).Count     ' return image count
            Index = Index + 1                   ' increment for next pass
            ExportImageList = True              ' return
        End If
    End With

End Function

Public Function ImportImageList(imgArray() As Byte, imgCount As Long) As Boolean

    ' UC use only. Called during its ReadProperties event
    ' Routine creates each imagelist from a PNG
    ' ExportImageList is the complimentary function
    
    ' imgArray :: a PNG file
    ' imgCount:: number of images in the imagelist

    ' Index begins at one and continues until function returns false
    
    If imgCount = 0& Then Exit Function      ' invalid parameter values
    Dim Ptr As Long, Index As Long
    
    If cLists Is Nothing Then Set cLists = New Collection
    
    Index = cLists.Count + 1                ' next imagelist to add
    ReDim Preserve m_Lists(1 To Index)      ' resize array
    
    ' create a new DIB section & then ask it to convert the PNG to 32bpp
    Set m_Lists(Index).Host = New c32bppDIB
    If m_Lists(Index).Host.LoadPicture_Stream(imgArray) = False Then
        
        ' failure and shouldn't happen. Rollback -- loss of data has occurred
        If Index = 0 Then
            Erase m_Lists
        Else
            Set m_Lists(Index).Host = Nothing
            ReDim Preserve m_Lists(1 To Index - 1)
        End If
        Exit Function
    End If
    
    ' ok, imagelist was created, now fill in the other properties
    With m_Lists(Index)
        .Count = imgCount                   ' number images in list
        .Width = .Host.Width \ .Count       ' width of each image
        cLists.Add Index, .Width & "x" & .Host.Height    ' add key lookup
    End With
    ImportImageList = True

End Function

Public Function ExportImageItemData(Index As Long, datArray() As Byte) As Boolean

    ' UC use only. Called during its WriteProperties event
    ' Routine serializes each imagelist item into 1024 byte chunks
    ' ImportImageItemData is the complimentary function
    
    ' on return, Index will be incremented for next set of items
    ' datArray :: will contain key, tag, index & slot in byte format
    '   max of 1024 bytes or full item if a single item > 1024 bytes
    
    ' Function kept calling until it returns false
    
    Dim I As Long, newIndex As Long
    Dim runningCount As Long, usableCount As Long
    Dim Ptr As Long
    
    Const fixedDat As Long = 8  ' slot & ListIndex (2 Longs * 4 bytes)
    
    If Index > m_ItemCount Or Index < 1 Then Exit Function  ' invalid parameter value
    
    ' 2 Loops, one to sum values, the other to write values
    
    ' Loop 1. See how many items we can add to a 1024 byte chunk
    For I = Index To m_ItemCount
        ' get length of Key & Tag values (allow unicode)
        runningCount = Len(m_Items(I).Tag) * 2& + 4&
        runningCount = runningCount + Len(m_Items(I).Key) * 2& + 4& + fixedDat
        ' ^^ also added was the 8 byte constant for slot & ListIndex
        If runningCount + usableCount > 1024 Then
            ' we exceeded our 1024 chunk...
            If usableCount = 0 Then         ' but no previous items were included
                usableCount = runningCount  ' so this is a big item (very long Key/Tag)
                newIndex = I + 1            ' This serialize will be only one item
            Else
                newIndex = I                ' Index to be included into this chunk
            End If
            Exit For
        Else
            usableCount = usableCount + runningCount    ' keep counting
        End If
    Next
    ' if we got all the way thru without Exit:For, then all items can fit in 1024 chunk
    If I > m_ItemCount Then newIndex = m_ItemCount
    
    ' size our chunk, may be less than 1024
    ReDim datArray(0 To usableCount - 1)
    
    ' Loop 2. Write the values into the chunk
    For I = Index To newIndex
        If m_Items(I).Key = vbNullString Then
            CopyMemory datArray(Ptr), 0&, 4&    ' null key
            runningCount = 0&
        Else
            ' here we will copy key directly from memory (unicode format)
            runningCount = Len(m_Items(I).Key) * 2
            CopyMemory datArray(Ptr), runningCount, 4&
            CopyMemory datArray(Ptr + 4), ByVal StrPtr(m_Items(I).Key), runningCount
        End If
        Ptr = runningCount + 4& + Ptr           ' increment pointer
        If m_Items(I).Tag = vbNullString Then
            CopyMemory datArray(Ptr), 0&, 4&    ' null key
            runningCount = 0&
        Else
            ' here we will copy tag directly from memory (unicode format)
            runningCount = Len(m_Items(I).Tag) * 2
            CopyMemory datArray(Ptr), runningCount, 4&
            CopyMemory datArray(Ptr + 4), ByVal StrPtr(m_Items(I).Tag), runningCount
        End If
        Ptr = runningCount + 4& + Ptr           ' increment pointer
        ' copy the non-pointer data: slot, listindex
        CopyMemory datArray(Ptr), ByVal VarPtr(m_Items(I)), fixedDat
        Ptr = Ptr + fixedDat                    ' increment pointer
    Next
    Index = newIndex + 1                        ' set next item for next pass
    ExportImageItemData = True
    
End Function

Public Function ImportImageItemData(datArray() As Byte) As Boolean

    ' UC use only. Called during its ReadProperties event
    ' Routine creates imagelist item data from passed array
    ' ExportImageItemData is the complimentary function
    
    ' datArray :: will contain key, tag, index, slot in byte format
    '   max of 1024 bytes or full item if a single item > 1024 bytes
    
    ' Function kept calling until it returns false
    
    Dim Ptr As Long
    Dim aValue As Long
    
    Const fixedDat As Long = 8  ' slot & ListIndex
    
    Ptr = LBound(datArray)
    ReDim Preserve m_Items(1 To m_ItemCount + 10)   ' add 10 extra vs redim'ing each pass
    
    Do Until Ptr > UBound(datArray)                 ' process entire array
        ' validate we don't need to resize our array again
        If m_ItemCount = UBound(m_Items) Then ReDim Preserve m_Items(1 To m_ItemCount + 10)
        m_ItemCount = m_ItemCount + 1               ' index to next imagelist item
        
        CopyMemory aValue, datArray(Ptr), 4& ' length of key
        If aValue > 0 Then
            m_Items(m_ItemCount).Key = String(aValue \ 2, 0)  ' create string with aValue bytes
            ' copy our array data directly to the String memory address
            CopyMemory ByVal StrPtr(m_Items(m_ItemCount).Key), datArray(Ptr + 4), aValue
            ' add the key to our keyed collection
            cItems.Add m_ItemCount, m_Items(m_ItemCount).Key
        End If
        Ptr = Ptr + aValue + 4              ' increment pointer
        CopyMemory aValue, datArray(Ptr), 4& ' length of tag
        If aValue > 0 Then                  ' same comments as above
            m_Items(m_ItemCount).Tag = String(aValue \ 2, 0)
            CopyMemory ByVal StrPtr(m_Items(m_ItemCount).Tag), datArray(Ptr + 4), aValue
        End If
        Ptr = Ptr + aValue + 4              ' increment pointer
        ' now copy the slot & listindex to our imagelist item
        CopyMemory ByVal VarPtr(m_Items(m_ItemCount)), datArray(Ptr), fixedDat
        Ptr = Ptr + fixedDat                ' increment for next imagelist item
        
    Loop
    
    ReDim Preserve m_Items(1 To m_ItemCount)    ' done, resize to exact size
    ImportImageItemData = True
    
End Function

Public Function ExportMiscProperties(datArray() As Byte) As Boolean

    ' function exports additional properties as needed
    ' Allows expandability without having to rewrite the uc's read/write properties code
    ' Returns false if there are not properties to export
    
    datArray() = ppgProps
    ' in future we can add additional values as long as the complimentary function recognizes them
    ExportMiscProperties = True

End Function

Public Function ImportMiscProperites(datArray() As Byte) As Boolean

    ppgProps = datArray()
    ImportMiscProperites = True
End Function

Private Sub Class_Initialize()
    Set cItems = New Collection
End Sub


