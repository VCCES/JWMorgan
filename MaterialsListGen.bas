Attribute VB_Name = "MaterialsListGen"
Option Explicit

Sub MaterialsListGen(b As clsBuilding)
Dim Qty As Integer
Dim rShape As String
Dim pShape As String
Dim pType As String
Dim rColor As String
Dim bLength As Double
Dim bWidth As Double
Dim bHeight As Double
Dim rPitch As String    ''' Roof Pitch Rise
Dim RafterLength As Double
Dim s2RafterSheetLength As Double
Dim s4RafterSheetLength As Double
Dim RoofPitchHypot As Double    ''' Inches Rise per Ft Roof Span
Dim s2EaveOverhang As Double
Dim s4EaveOverhang As Double
Dim e1GableOverhang As Double
Dim e3GableOverhang As Double
'new roof panel calculations
Dim s2RoofPanels As Collection
Dim s4RoofPanels As Collection
Dim e1GableExtensionPanels As Collection
Dim e3GableExtensionPanels As Collection
Dim RoofLength As Integer
Dim RoofPanel As clsPanel
'Standard Overhang
Dim StandardEaveOverhang As Double
Dim Feet As Single
Dim Inches As Double
Dim InchFraction As Double
Dim Undercut As Double
'''''''''' ridge cap
Dim RidgeCapQty As Integer
Dim PitchString As String
''' sidewall
Dim wShape As String
Dim wType As String
Dim wColor As String
''' High sidewall height
Dim HighSideEaveHeight As Double
'sidewall panels
Dim s2SidewallPanels As New Collection
Dim s4SidewallPanels As New Collection
''' endwalls
Dim e1EndwallPanels As New Collection
Dim e3EndwallPanels As New Collection
Dim EndwallPanelCount As Double
Dim e1PanelQty As Integer
Dim e1PanelLength As String
Dim e3PanelQty As Integer
Dim e3PanelLength As String
Dim PanelNumber As Integer
Dim pLength As Double
Dim ePanel As clsPanel
Dim p As Integer
Dim MaxHeight As Integer
'wall panel class for vendor material list
Dim WallPanel As clsPanel
'''''' Rake Trim
Dim RakeTrimPieces As Collection
Dim NetRafterLength As Double
Dim RakeTrimColor As String
Dim TrimPiece As clsTrim
'''''' Eave Trim
Dim EaveTrimPieces As Collection
Dim s2EaveTrimLength As Double
Dim s4EaveTrimLength As Double
Dim EaveTrimColor As String
'''''' Outside Corner Trim
Dim OutsideCornerTrimPieces As Collection
Dim NetCornerLength As Double
Dim OutsideCornerTrimColor As String
'''''' Base Trim
Dim BaseTrimPieces As Collection
Dim NetBaseTrimLength As Double
Dim NetPDoorWidth As Double '' For net personelle door
Dim NetOHDoorWidth As Double '' For net overhead door
Dim BaseTrimColor As String
'''''Wainscot Trim
Dim WainscotTrimPieces As Collection
Dim TempDoorWidth As Double ''For Wainscot Trim
Dim NetWainscotTrimLength As Double
'''''' Framed Openings
Dim FOCell As Range

'''''' Gutters & Downspouts
Dim GutterPieces As Collection
Dim Gutters As Boolean
Dim NetGutterLength As Double
Dim GutterEndCapQty As Integer
Dim GutterStrapQty As Integer
Dim GutterPiece As clsTrim  'trim class since available in the same sizes
Dim GutterColor As String
Dim DownspoutColor As String
Dim DownspoutQty As Integer
Dim DownspoutPieces As Collection
Dim DownspoutPiece As clsTrim 'trim class since available in the same sizes
Dim RemainingHeight As Double
Dim DownspoutStrapQty As Integer
Dim h As Integer
Dim PopRivitQty As Integer
'Bays
Dim BayCount As Integer
' Translucent Wall Panels & Skylights
Dim SkylightPanelQty As Integer
Dim SkylightPanel As clsPanel
'''''' Soffits
Dim e1GableOverhangSoffit As Boolean
Dim e1GableExtensionSoffit As Boolean
Dim s2EaveExtensionSoffit As Boolean
Dim s2EaveOverhangSoffit As Boolean
Dim e3GableOverhangSoffit As Boolean
Dim e3GableExtensionSoffit As Boolean
Dim s4EaveOverhangSoffit As Boolean
Dim s4EaveExtensionSoffit As Boolean
Dim NetOutsideAngleLength As Double
Dim SoffitPanel As clsPanel
Dim SoffitTrim As clsTrim
Dim SoffitPiece As Variant
Dim SoffitQty As Integer
''' Extensions
Dim s2EaveExtension As Double
Dim s4EaveExtension As Double
Dim e1GableExtension As Double
Dim e3GableExtension As Double
Dim PanelQty As Integer
Dim e1ExtensionPanels As Collection
Dim s2ExtensionPanels As Collection
Dim e3ExtensionPanels As Collection
Dim s4ExtensionPanels As Collection
Dim ExtensionPanel As clsPanel
''' Overhangs
Dim e1GableOverhangSection As Boolean
Dim s2EaveOverhangSection As Boolean
Dim e3GableOverhangSection As Boolean
Dim s4EaveOverhangSection As Boolean
Dim e1GableExtensionSection As Boolean
Dim s2EaveExtensionSection As Boolean
Dim e3GableExtensionSection As Boolean
Dim s4EaveExtensionSection As Boolean
'overhang or extension collections
' soffit collections
Dim e1SoffitPanels As Collection
Dim e1SoffitTrim As Collection
Dim e1ExtensionSoffitTrim As Collection
Dim s2SoffitPanels As Collection
Dim s2SoffitTrim As Collection
Dim s2ExtensionSoffitTrim As Collection
Dim e3SoffitPanels As Collection
Dim e3SoffitTrim As Collection
Dim e3ExtensionSoffitTrim As Collection
Dim s4SoffitPanels As Collection
Dim s4SoffitTrim As Collection
Dim s4ExtensionSoffitTrim As Collection
'2x8 inside angle
Dim e1InsideAngleTrim As Collection
Dim e3InsideAngleTrim As Collection
Dim NetInsideAngleLength As Double
''' Fasteners
Dim rTekScrewQty As Integer
Dim rLapScrewQty As Integer
Dim wTekScrewQty As Integer
Dim wLapScrewQty As Integer
Dim rPurlins As Integer
Dim pTypeCount As Integer
Dim rOverlaps As Integer
Dim sOverlaps As Integer
Dim eOverlaps As Integer
Dim TrimScrewQty As Integer
Dim NetRakeTrimLength As Integer
Dim SoffitScrewQty As Integer
Dim SoffitScrewColor As String
Dim Screw As clsFastener
Dim TrimScrews As Collection
Dim SoffitScrews As Collection
'Liner Panels
Dim e1LinerPanels As New Collection
Dim e3LinerPanels As New Collection
Dim s2LinerPanels As New Collection
Dim s4LinerPanels As New Collection
Dim RoofLinerPanels As New Collection
Dim LinerPanelsSection As Boolean
'clsPanel
Dim Panel As clsPanel
''' Miscellaneous
Dim ButylTapeQty As Integer
Dim InsideClosureQty As Integer
Dim OutsideClosureQty As Integer
''' Vendor Materials List
Dim PanelCollection As Collection
Dim TrimCollection As Collection
Dim MiscCollection As Collection
Dim item As clsMiscItem

'Misc Variables
Dim MatSht As Worksheet
Dim n As Integer
Dim WriteCell As Range




''''''''' Generate Roofing Section

''' Read Information
With EstSht
    'building width, building length, roof pitch
    bWidth = .Range("Building_Width").Value
    bHeight = .Range("Building_Height").Value
    bLength = .Range("Building_Length").Value
    rPitch = .Range("Roof_Pitch").Value
    'single slope or gable
    rShape = .Range("Roof_Shape").Value
    'roof panel info
    pShape = .Range("Roof_pShape").Value
    pType = .Range("Roof_pType").Value
    rColor = .Range("Roof_Color").Value
    ' wall panel info
    wShape = .Range("Wall_pShape").Value
    wType = .Range("Wall_pType").Value
    wColor = .Range("Wall_Color").Value
    
    'check for invalid building height
    If bHeight > 80 Then
        MsgBox "Buildings cannot be taller than 80'. Please correct the data before generating a materials list.", vbExclamation, "Building Height Error"
        End
    ElseIf rShape = "Single Slope" Then
        If (bHeight + ((bWidth * rPitch) / 12)) > 100 Then
            MsgBox "The high side eave cannot be greater than 100'. Please correct the data before generating a materials list.", vbExclamation, "High Side Eave Error"
            End
        End If
    End If
    
    ''' overhang info
    'convert to inches
    e1GableOverhang = .Range("e1_GableOverhang").Value * 12
    s2EaveOverhang = .Range("s2_EaveOverhang").Value * 12
    e3GableOverhang = .Range("e3_GableOverhang").Value * 12
    s4EaveOverhang = .Range("s4_EaveOverhang").Value * 12
    
    ''' Extensions
    e1GableExtension = .Range("e1_GableExtension").Value * 12
    s2EaveExtension = .Range("s2_EaveExtension").Value * 12
    e3GableExtension = .Range("e3_GableExtension").Value * 12
    s4EaveExtension = .Range("s4_EaveExtension").Value * 12
    
    ''' Trim
    RakeTrimColor = .Range("Rake_tColor").Value
    EaveTrimColor = .Range("Eave_tColor").Value
    OutsideCornerTrimColor = .Range("OutsideCorner_tColor").Value
    BaseTrimColor = .Range("Base_tColor").Value
    
    ''' Gutters
    If .Range("GutterAndDownspouts").Value = "Yes" Then Gutters = True
    GutterColor = .Range("GutterColor").Value
    DownspoutColor = .Range("DownspoutColor").Value
    
    ''' Soffits
    'check if soffits
    If .Range("e1_GableOverhangSoffit").Value = "Yes" Then e1GableOverhangSoffit = True
    If .Range("e1_GableExtensionSoffit").Value = "Yes" Then e1GableExtensionSoffit = True
    If .Range("s2_EaveOverhangSoffit").Value = "Yes" Then s2EaveOverhangSoffit = True
    If .Range("s2_EaveExtensionSoffit").Value = "Yes" Then s2EaveExtensionSoffit = True
    If .Range("e3_GableOverhangSoffit").Value = "Yes" Then e3GableOverhangSoffit = True
    If .Range("e3_GableExtensionSoffit").Value = "Yes" Then e3GableExtensionSoffit = True
    If .Range("s4_EaveOverhangSoffit").Value = "Yes" Then s4EaveOverhangSoffit = True
    If .Range("s4_EaveExtensionSoffit").Value = "Yes" Then s4EaveExtensionSoffit = True
    
    ''' Overhangs
    If .Range("e1_GableOverhang").Value > 0 And e1GableOverhangSoffit = True Then e1GableOverhangSection = True
    If .Range("s2_EaveOverhang").Value > 0 And s2EaveOverhangSoffit = True Then s2EaveOverhangSection = True
    If .Range("e3_GableOverhang").Value > 0 And e3GableOverhangSoffit = True Then e3GableOverhangSection = True
    If .Range("s4_EaveOverhang").Value > 0 And s4EaveOverhangSoffit = True Then s4EaveOverhangSection = True
    ''' Extensions
    If .Range("e1_GableExtension").Value > 0 Then e1GableExtensionSection = True
    If .Range("s2_EaveExtension").Value > 0 Then s2EaveExtensionSection = True
    If .Range("e3_GableExtension").Value > 0 Then e3GableExtensionSection = True
    If .Range("s4_EaveExtension").Value > 0 Then s4EaveExtensionSection = True
    
    'add in standard overhangs
    s2EaveOverhang = s2EaveOverhang + 4.25
    '''s4
    'always additional 4.25 overhang for gable or if an s4 eave extension
    If rShape = "Gable" Or s4EaveExtensionSection = True Then s4EaveOverhang = s4EaveOverhang + 4.25
    'for single slope, no additional 4.25" s4 overhang as long as there's no extension
    
    '' building class setup '''
    b.bHeight = bHeight
    b.bLength = bLength
    b.bWidth = bWidth
    b.e1Overhang = e1GableOverhang
    b.e3Overhang = e3GableOverhang
    b.s2Overhang = s2EaveOverhang
    b.s4Overhang = s4EaveOverhang
    b.e1Extension = e1GableExtension
    b.e3Extension = e3GableExtension
    b.s2Extension = s2EaveExtension
    b.s4Extension = s4EaveExtension
    b.rPitch = rPitch
    b.rShape = rShape
    b.Gutters = Gutters
    b.wPanelShape = wShape
    b.wPanelColor = wColor
    b.wPanelType = wType
    b.rPanelShape = pShape
    b.rPanelType = pType
    b.rPanelColor = rColor
    b.RakeTrimColor = RakeTrimColor
    b.OutsideCornerTrimColor = OutsideCornerTrimColor
    'check for base trim
    If BaseTrimColor = "None" Or BaseTrimColor = "" Then
        b.BaseTrim = False
    Else
        b.BaseTrim = True
    End If
    'set panel overage along building length
    b.bLengthRoofPanelOverage = (Application.WorksheetFunction.RoundUp(((bLength * 12) + e1GableOverhang + e3GableOverhang) / (3 * 12), 0) * 3 * 12) - ((bLength * 12) + e1GableOverhang + e3GableOverhang)
    'soffit booleans
    If e1GableOverhangSoffit = True Then b.e1GableOverhangSoffit = True
    If e3GableOverhangSoffit = True Then b.e3GableOverhangSoffit = True
    If s2EaveOverhangSoffit = True Then b.s2EaveOverhangSoffit = True
    If s4EaveOverhangSoffit = True Then b.s4EaveOverhangSoffit = True
    If e1GableExtensionSoffit = True Then b.e1GableExtensionSoffit = True
    If e3GableExtensionSoffit = True Then b.e3GableExtensionSoffit = True
    If s2EaveExtensionSoffit = True Then b.s2EaveExtensionSoffit = True
    If s4EaveExtensionSoffit = True Then b.s4EaveExtensionSoffit = True
    
    'eave extension pitches
    With EstSht
        's2 eave extension
        If .Range("s2_EaveExtensionPitch").Value = "Match Roof" Then
            b.s2ExtensionPitch = rPitch
        Else
            b.s2ExtensionPitch = .Range("s2_EaveExtensionPitch").Value
        End If
        's4 eave extension
        If .Range("s4_EaveExtensionPitch").Value = "Match Roof" Then
            b.s4ExtensionPitch = rPitch
        Else
            b.s4ExtensionPitch = .Range("s4_EaveExtensionPitch").Value
        End If
    End With
            
    ''' Liner Panels Section
    If b.LinerPanels("e1") = "8'" Or b.LinerPanels("e1") = "Full Height" _
    Or b.LinerPanels("e3") = "8'" Or b.LinerPanels("e3") = "Full Height" _
    Or b.LinerPanels("s2") = "8'" Or b.LinerPanels("s2") = "Full Height" _
    Or b.LinerPanels("s4") = "8'" Or b.LinerPanels("s4") = "Full Height" Then LinerPanelsSection = True
    
    ''' extension overhangs
    If e1GableOverhangSection = True And e1GableExtensionSection = True Then
        b.e1ExtensionOverhang = b.e1Overhang
        b.e1Overhang = 0
    End If
    If e3GableOverhangSection = True And e3GableExtensionSection = True Then
        b.e3ExtensionOverhang = b.e3Overhang
        b.e3Overhang = 0
    End If
    If s2EaveOverhangSection = True And s2EaveExtensionSection = True Then
        b.s2ExtensionOverhang = b.s2Overhang
        b.s2Overhang = 0
    End If
    If s4EaveOverhangSection = True And s4EaveExtensionSection = True Then
        b.s4ExtensionOverhang = b.s4Overhang
        b.s4Overhang = 0
    End If
        
    
    ''' Framed Openings
    'Personell Doors
    For Each FOCell In Range(.Range("pDoorCell1"), .Range("pDoorCell12"))
        'if cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'add size to perimeter
            If FOCell.offset(0, 1).Value = "3070" Then
                NetPDoorWidth = NetPDoorWidth + 3
            ElseIf FOCell.offset(0, 1).Value = "4070" Then
                NetPDoorWidth = NetPDoorWidth + 4
            End If
        End If
    Next FOCell
    'Overhead Doors
    For Each FOCell In Range(.Range("OHDoorCell1"), .Range("OHDoorCell12"))
        'if cell isn't hidden, door width is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'add size to perimeter
            NetOHDoorWidth = NetOHDoorWidth + FOCell.offset(0, 1).Value
        End If
    Next FOCell
    
    ''' Bays
    BayCount = .Range("BayNum").Value
    
    
End With

'standard undercut
Undercut = 4.25

'roof pitch string for product names
PitchString = rPitch & ":12"


    

'''check for necessary roof information
If bWidth = 0 Or rPitch = 0 Or rShape = "" Or pShape = "" Or pType = "" Or _
pType = "" Or rColor = "" Then
    GoTo MissingRoofData
End If

''''' Roof Pitch Hypotenuse (inches per ft width)
RoofPitchHypot = Sqr((rPitch) ^ 2 + (12) ^ 2)

'roof length
RoofLength = bLength + (b.e1Overhang / 12) + (b.e3Overhang / 12)
'RoofLength = bLength + (e1GableOverhang / 12) + (e3GableOverhang / 12)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' For Gable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If rShape = "Gable" Then
    ''''''''''''' Panel Length '''''''''''''''''
    
    'normal roof rafter length (inches)
    RafterLength = (bWidth / 2) * RoofPitchHypot
    b.RafterLength = RafterLength
    'B.RafterLength = RafterLength - Undercut
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' sidewall 2 roof panels'''''
    '''''''''''' generate sidewall 2 roof panels
    Set s2RoofPanels = New Collection
    '''''' check for extension and overhang
    If b.s2ExtensionOverhang > 0 Then
        'extension and overhang case: subtract standard overhang
        s2RafterSheetLength = RafterLength + 4.25 - Undercut
        b.s2RafterSheetLength = s2RafterSheetLength
        b.s2ExtensionOverhang = b.s2ExtensionOverhang - 4.25
        'update extension/overhang stored length
        'b.s2Extension = b.s2Extension + b.s2Overhang - 4.25
        s2EaveOverhang = 0
        'b.s2Overhang = 0
        Call RoofPanelGen(s2RoofPanels, b.s2RafterSheetLength, 4.25, RoofLength, rShape)
    Else
        ''normal overhang handling
        'add undercut to rafter length
        s2RafterSheetLength = RafterLength + s2EaveOverhang - Undercut
        b.s2RafterSheetLength = s2RafterSheetLength
        Call RoofPanelGen(s2RoofPanels, b.s2RafterSheetLength, b.s2Overhang, RoofLength, rShape)
    End If
    'add qualities
    For Each RoofPanel In s2RoofPanels
        RoofPanel.PanelShape = pShape
        RoofPanel.PanelType = pType
        RoofPanel.PanelColor = rColor
    Next RoofPanel
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' sidewall 4 roof panels'''''
    '''''''''''' generate sidewall 4 roof panels
    Set s4RoofPanels = New Collection
    '''''' check for extension and overhang
    If b.s4ExtensionOverhang > 0 Then
        'extension and overhang case: Just add in the standard overhang
        s4RafterSheetLength = RafterLength + 4.25 - Undercut
        b.s4RafterSheetLength = s4RafterSheetLength
        b.s4ExtensionOverhang = b.s4ExtensionOverhang - 4.25
        'update extension/overhang stored length
        'b.s4Extension = b.s4Extension + b.s4Overhang - 4.25
        s4EaveOverhang = 0
        'b.s4Overhang = 0
        Call RoofPanelGen(s4RoofPanels, b.s4RafterSheetLength, 4.25, RoofLength, rShape)
    Else
        ''normal overhang handling
        'add undercut to rafter length
        s4RafterSheetLength = RafterLength + s4EaveOverhang - Undercut
        b.s4RafterSheetLength = s4RafterSheetLength
        Call RoofPanelGen(s4RoofPanels, b.s4RafterSheetLength, b.s4Overhang, RoofLength, rShape)
    End If
    'add qualities
    For Each RoofPanel In s4RoofPanels
        RoofPanel.PanelShape = pShape
        RoofPanel.PanelType = pType
        RoofPanel.PanelColor = rColor
    Next RoofPanel
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate ridge cap qty
    RidgeCapQty = Application.WorksheetFunction.RoundUp(((bLength + (b.e1Overhang / 12) + (b.e3Overhang / 12) + (b.e1Extension / 12) + (b.e3Extension / 12) + (b.e1ExtensionOverhang / 12) + (b.e3ExtensionOverhang / 12)) / 3), 0)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sidewall Panels
    Call SidewallPanelGen(s2SidewallPanels, "s2", b)
    Call SidewallPanelGen(s4SidewallPanels, "s4", b)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Endwall Panels
    Call EndwallPanelGen(e1EndwallPanels, "e1", b)
    Call EndwallPanelGen(e3EndwallPanels, "e3", b)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' For Single Slope
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf rShape = "Single Slope" Then
    ''''''''''''' Panel Length '''''''''''''''''
    
    Set s2RoofPanels = New Collection
    'normal roof rafter length (in)
    RafterLength = bWidth * RoofPitchHypot
    b.RafterLength = RafterLength
    
    
    ''' check case: extension and overhangs
    Select Case True
    ''' extension and overhang on both s2 and s4
    Case b.s2ExtensionOverhang > 0 And b.s4ExtensionOverhang > 0
        s2RafterSheetLength = RafterLength + 4.25 + 4.25
        b.s2RafterSheetLength = s2RafterSheetLength
        b.s2ExtensionOverhang = b.s2ExtensionOverhang - 4.25
        b.s4ExtensionOverhang = b.s4ExtensionOverhang - 4.25
        Call RoofPanelGen(s2RoofPanels, s2RafterSheetLength, 4.25, RoofLength, rShape)
    ''' extension and overhang on just s2
    Case b.s2Overhang > 4.25 And b.s2Extension > 0
        s2RafterSheetLength = RafterLength + 4.25 + s4EaveOverhang
        b.s2RafterSheetLength = s2RafterSheetLength
        b.s2ExtensionOverhang = b.s2ExtensionOverhang - 4.25
        Call RoofPanelGen(s2RoofPanels, s2RafterSheetLength, 4.25, RoofLength, rShape)
    ''' extension and overhang on just s4
    Case b.s4Overhang > 4.25 And b.s4Extension > 0
        s2RafterSheetLength = RafterLength + 4.25 + s2EaveOverhang
        b.s2RafterSheetLength = s2RafterSheetLength
        b.s4ExtensionOverhang = b.s4ExtensionOverhang - 4.25
        Call RoofPanelGen(s2RoofPanels, s2RafterSheetLength, s2EaveOverhang, RoofLength, rShape)
    ''' normal handling: no extension and overhang added
    Case Else
        s2RafterSheetLength = RafterLength + s2EaveOverhang + s4EaveOverhang
        b.s2RafterSheetLength = s2RafterSheetLength
        Call RoofPanelGen(s2RoofPanels, s2RafterSheetLength, s2EaveOverhang, RoofLength, rShape)
    End Select
    'add qualities
    For Each RoofPanel In s2RoofPanels
        RoofPanel.PanelShape = pShape
        RoofPanel.PanelType = pType
        RoofPanel.PanelColor = rColor
    Next RoofPanel
    'blank sidewall 4 collection
    Set s4RoofPanels = New Collection

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sidewall Panels

   Call SidewallPanelGen(s2SidewallPanels, "s2", b)
   Call SidewallPanelGen(s4SidewallPanels, "s4", b)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Endwall Panels
    Call EndwallPanelGen(e1EndwallPanels, "e1", b)
    Call EndwallPanelGen(e3EndwallPanels, "e3", b)
    
End If

''''''''''''''''''''''''''''''''''''''''''Liner Panels
Call LinerPanelGen(e1LinerPanels, b, "e1")
Call LinerPanelGen(e3LinerPanels, b, "e3")
Call LinerPanelGen(s2LinerPanels, b, "s2")
Call LinerPanelGen(s4LinerPanels, b, "s4")
Call LinerPanelGen(RoofLinerPanels, b, "Roof")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Trim
''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Rake Trim (For Either roof Type)
''''
Set RakeTrimPieces = New Collection
'calculate net rafter length (without factoring in undercut)
If rShape = "Single Slope" Then
    NetRafterLength = b.RafterLength
ElseIf rShape = "Gable" Then
    NetRafterLength = b.RafterLength * 2
End If
'add extension/overhang
NetRafterLength = NetRafterLength + b.s2ExtensionRafterLength + b.s4ExtensionRafterLength
'two endwalls, two trim lengths
NetRafterLength = NetRafterLength * 2
'pass to calc
Call TrimPieceCalc(RakeTrimPieces, NetRafterLength, "Rake", , , b)
''''
'add qualities
For Each TrimPiece In RakeTrimPieces
    TrimPiece.tShape = pShape   ''' rake trim is always roof panel shape
    TrimPiece.Color = RakeTrimColor
    'increase 20'4" pieces to 21'
    If TrimPiece.tLength = 244 Then
        TrimPiece.tLength = 21 * 12
        TrimPiece.tMeasurement = ImperialMeasurementFormat(TrimPiece.tLength)
    End If
Next TrimPiece
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Eave Trim
''''
Set EaveTrimPieces = New Collection
'''''''''''''''''''''''''' Sidewall 2
''' Outside Eave Trim (Along outside of Eave)
'start with building length
s2EaveTrimLength = bLength * 12
'add endwall overhangs or extensions
s2EaveTrimLength = s2EaveTrimLength + b.e1Extension + b.e1ExtensionOverhang + b.e3Extension + b.e3ExtensionOverhang + b.e3Extension + b.e1Overhang
''' Inside Eave Trim (Where eave meets sidewalls)
'' Check if inside eave trim is needed
'First check if there's an s2 extension/overhang
If (b.s2Extension <> 0) Or (b.s2Overhang <> 4.25) Then ''' s2 eave extension/overhang present                                               '''
    'Add inside Trim if no eave soffit
    If s2EaveOverhangSoffit = False And s2EaveExtensionSoffit = False Then
        ' add additional trim along building length
        s2EaveTrimLength = s2EaveTrimLength + (bLength * 12)
    End If
End If

'''''''''''''''''''''''''' Sidewall 4, generate trim collections
''' Outside Eave Trim (Along outside of Eave)
'start with building length
s4EaveTrimLength = bLength * 12
'add endwall overhangs or extensions
s4EaveTrimLength = s4EaveTrimLength + b.e1Extension + b.e1ExtensionOverhang + b.e3Extension + b.e3ExtensionOverhang + b.e3Extension + b.e1Overhang
''' Inside Eave Trim If Needed (Where eave meets sidewalls)
'Additional condition due 4.25 s4 standard overhang on gable and 0 s4 standard overhang on single slope
If (b.s4Extension <> 0) Or ((b.s4Overhang <> 4.25) And (b.s4Overhang <> 0)) Then   ''' s4 eave extension/overhang present                                               '''
    'Add inside Trim if no eave soffit
    If s4EaveExtensionSoffit = False And s4EaveOverhangSoffit = False Then
        ' add additional trim along building length
        s4EaveTrimLength = s4EaveTrimLength + (bLength * 12)
    End If
End If
    
If rShape = "Gable" Then
    ''''generate trim piece collection
    Call TrimPieceCalc(EaveTrimPieces, s2EaveTrimLength + s4EaveTrimLength, "Short Eave", PitchString)
ElseIf rShape = "Single Slope" Then
    'seperate high side and short side eave trim collections
    Call TrimPieceCalc(EaveTrimPieces, s2EaveTrimLength, "Short Eave", PitchString)
    Call TrimPieceCalc(EaveTrimPieces, s4EaveTrimLength, "High Eave", PitchString)
End If

'add qualities
For Each TrimPiece In EaveTrimPieces
    TrimPiece.tShape = "R-Loc"
    TrimPiece.Color = EaveTrimColor
Next TrimPiece
''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Outside Corner Trim
''''
Set OutsideCornerTrimPieces = New Collection
If b.rShape = "Gable" Then
    'assume complete, exclude intersections if needed
    NetCornerLength = b.bHeight * 4 * 12
    If b.WallStatus("e1") <> "Include" And b.WallStatus("s2") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("e1") <> "Include" And b.WallStatus("s4") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("e3") <> "Include" And b.WallStatus("s2") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("e3") <> "Include" And b.WallStatus("s4") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
ElseIf rShape = "Single Slope" Then
    'sidewall 2 corners + s4 corners
    NetCornerLength = (b.bHeight * 12 * 2) + (b.HighSideEaveHeight * 2)
    'exclude where needed
    If b.WallStatus("s2") <> "Include" And b.WallStatus("e1") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("s2") <> "Include" And b.WallStatus("e3") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("s4") <> "Include" And b.WallStatus("e1") <> "Include" Then NetCornerLength = NetCornerLength - b.HighSideEaveHeight
    If b.WallStatus("s4") <> "Include" And b.WallStatus("e3") <> "Include" Then NetCornerLength = NetCornerLength - b.HighSideEaveHeight
End If
'generate trim collection
Call TrimPieceCalc(OutsideCornerTrimPieces, NetCornerLength, "Outside Corner", , , b)
'add qualities
For Each TrimPiece In OutsideCornerTrimPieces
    TrimPiece.tShape = "R-Loc"
    TrimPiece.Color = OutsideCornerTrimColor
Next TrimPiece
''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Base Trim
''''
If b.BaseTrim = True Then
    Set BaseTrimPieces = New Collection
    'Perimeter
    If b.WallStatus("e1") = "Include" Then NetBaseTrimLength = b.bWidth
    If b.WallStatus("e3") = "Include" Then NetBaseTrimLength = NetBaseTrimLength + b.bWidth
    If b.WallStatus("s2") = "Include" Then NetBaseTrimLength = NetBaseTrimLength + b.bLength
    If b.WallStatus("s4") = "Include" Then NetBaseTrimLength = NetBaseTrimLength + b.bLength
    'subtract width of OH doors, P doors
    NetBaseTrimLength = NetBaseTrimLength - (NetOHDoorWidth + NetPDoorWidth)
    'convert
    NetBaseTrimLength = NetBaseTrimLength * 12
    'generate trim collection
    Call TrimPieceCalc(BaseTrimPieces, NetBaseTrimLength, "Base")
    'add qualities
    For Each TrimPiece In BaseTrimPieces
        TrimPiece.tShape = "R-Loc"
        TrimPiece.Color = BaseTrimColor
    Next TrimPiece
End If
''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Wainscot Trim
''''
If EstSht.Range("Wainscot").Value = "Yes" Then
    Set WainscotTrimPieces = New Collection
    '''''''''''''''''''''''''''''''''''''''''Standard Wainscot Trim'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Endwall 1 - check if standard
    If b.Wainscot("e1") <> "None" And InStr(1, b.Wainscot("e1"), "Standard") Then
        'loop through Pdoors for door widths on Endwall 1
        For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
            'if cell isn't hidden, door size is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 2).Value = "Endwall 1" Then
                'add size to perimeter
                If FOCell.offset(0, 1).Value = "3070" Then
                    TempDoorWidth = TempDoorWidth + 3
                ElseIf FOCell.offset(0, 1).Value = "4070" Then
                    TempDoorWidth = TempDoorWidth + 4
                End If
            End If
        Next FOCell
        'loop through OHdoors for door widths on Endwall 1
        For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
            'if cell isn't hidden, door width is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 3).Value = "Endwall 1" Then
                'add size to perimeter
                TempDoorWidth = TempDoorWidth + FOCell.offset(0, 1).Value
            End If
        Next FOCell
        NetWainscotTrimLength = b.bWidth
    End If
    'Endwall 3 - check if standard
    If b.Wainscot("e3") <> "None" And InStr(1, b.Wainscot("e3"), "Standard") Then
        For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
            'if cell isn't hidden, door size is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 2).Value = "Endwall 3" Then
                'add size to perimeter
                If FOCell.offset(0, 1).Value = "3070" Then
                    TempDoorWidth = TempDoorWidth + 3
                ElseIf FOCell.offset(0, 1).Value = "4070" Then
                    TempDoorWidth = TempDoorWidth + 4
                End If
            End If
        Next FOCell
        For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
            'if cell isn't hidden, door width is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 3).Value = "Endwall 3" Then
                'add size to perimeter
                TempDoorWidth = TempDoorWidth + FOCell.offset(0, 1).Value
            End If
        Next FOCell
        NetWainscotTrimLength = b.bWidth + NetWainscotTrimLength
    End If
    'Sidewall 2 - check if standard
    If b.Wainscot("s2") <> "None" And InStr(1, b.Wainscot("s2"), "Standard") Then
        For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
            'if cell isn't hidden, door size is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 2).Value = "Sidewall 2" Then
                'add size to perimeter
                If FOCell.offset(0, 1).Value = "3070" Then
                    TempDoorWidth = TempDoorWidth + 3
                ElseIf FOCell.offset(0, 1).Value = "4070" Then
                    TempDoorWidth = TempDoorWidth + 4
                End If
            End If
        Next FOCell
        For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
            'if cell isn't hidden, door width is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 3).Value = "Sidewall 2" Then
                'add size to perimeter
                TempDoorWidth = TempDoorWidth + FOCell.offset(0, 1).Value
            End If
        Next FOCell
        NetWainscotTrimLength = b.bLength + NetWainscotTrimLength
    End If
    'Sidewall 4 - check if standard
    If b.Wainscot("s4") <> "None" And InStr(1, b.Wainscot("s4"), "Standard") Then
        For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
            'if cell isn't hidden, door size is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 2).Value = "Sidewall 4" Then
                'add size to perimeter
                If FOCell.offset(0, 1).Value = "3070" Then
                    TempDoorWidth = TempDoorWidth + 3
                ElseIf FOCell.offset(0, 1).Value = "4070" Then
                    TempDoorWidth = TempDoorWidth + 4
                End If
            End If
        Next FOCell
        For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
            'if cell isn't hidden, door width is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 3).Value = "Sidewall 4" Then
                'add size to perimeter
                TempDoorWidth = TempDoorWidth + FOCell.offset(0, 1).Value
            End If
        Next FOCell
        NetWainscotTrimLength = b.bLength + NetWainscotTrimLength
    End If
    'subtract width of OH doors and Pdoors where there is standard wainscot
    NetWainscotTrimLength = NetWainscotTrimLength - TempDoorWidth
    'convert
    NetWainscotTrimLength = NetWainscotTrimLength * 12
    'generate trim collection
    Call TrimPieceCalc(WainscotTrimPieces, NetWainscotTrimLength, "Standard Wainscot")
    'Reset Variables
    TempDoorWidth = 0
    NetWainscotTrimLength = 0
    '''''''''''''''''''''''''''''''''''''''''Masonry Wainscot Trim'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Endwall 1 - check if standard
    If b.Wainscot("e1") <> "None" And InStr(1, b.Wainscot("e1"), "Masonry") Then
        'loop through Pdoors for door widths on Endwall 1
        For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
            'if cell isn't hidden, door size is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 2).Value = "Endwall 1" Then
                'add size to perimeter
                If FOCell.offset(0, 1).Value = "3070" Then
                    TempDoorWidth = TempDoorWidth + 3
                ElseIf FOCell.offset(0, 1).Value = "4070" Then
                    TempDoorWidth = TempDoorWidth + 4
                End If
            End If
        Next FOCell
        'loop through OHdoors for door widths on Endwall 1
        For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
            'if cell isn't hidden, door width is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 3).Value = "Endwall 1" Then
                'add size to perimeter
                TempDoorWidth = TempDoorWidth + FOCell.offset(0, 1).Value
            End If
        Next FOCell
        NetWainscotTrimLength = b.bWidth
    End If
    'Endwall 3 - check if standard
    If b.Wainscot("e3") <> "None" And InStr(1, b.Wainscot("e3"), "Masonry") Then
        For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
            'if cell isn't hidden, door size is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 2).Value = "Endwall 3" Then
                'add size to perimeter
                If FOCell.offset(0, 1).Value = "3070" Then
                    TempDoorWidth = TempDoorWidth + 3
                ElseIf FOCell.offset(0, 1).Value = "4070" Then
                    TempDoorWidth = TempDoorWidth + 4
                End If
            End If
        Next FOCell
        For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
            'if cell isn't hidden, door width is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 3).Value = "Endwall 3" Then
                'add size to perimeter
                TempDoorWidth = TempDoorWidth + FOCell.offset(0, 1).Value
            End If
        Next FOCell
        NetWainscotTrimLength = b.bWidth + NetWainscotTrimLength
    End If
    'Sidewall 2 - check if standard
    If b.Wainscot("s2") <> "None" And InStr(1, b.Wainscot("s2"), "Masonry") Then
        For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
            'if cell isn't hidden, door size is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 2).Value = "Sidewall 2" Then
                'add size to perimeter
                If FOCell.offset(0, 1).Value = "3070" Then
                    TempDoorWidth = TempDoorWidth + 3
                ElseIf FOCell.offset(0, 1).Value = "4070" Then
                    TempDoorWidth = TempDoorWidth + 4
                End If
            End If
        Next FOCell
        For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
            'if cell isn't hidden, door width is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 3).Value = "Sidewall 2" Then
                'add size to perimeter
                TempDoorWidth = TempDoorWidth + FOCell.offset(0, 1).Value
            End If
        Next FOCell
        NetWainscotTrimLength = b.bLength + NetWainscotTrimLength
    End If
    'Sidewall 4 - check if standard
    If b.Wainscot("s4") <> "None" And InStr(1, b.Wainscot("s4"), "Masonry") Then
        For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
            'if cell isn't hidden, door size is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 2).Value = "Sidewall 4" Then
                'add size to perimeter
                If FOCell.offset(0, 1).Value = "3070" Then
                    TempDoorWidth = TempDoorWidth + 3
                ElseIf FOCell.offset(0, 1).Value = "4070" Then
                    TempDoorWidth = TempDoorWidth + 4
                End If
            End If
        Next FOCell
        For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
            'if cell isn't hidden, door width is entered
            If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" And FOCell.offset(0, 3).Value = "Sidewall 4" Then
                'add size to perimeter
                TempDoorWidth = TempDoorWidth + FOCell.offset(0, 1).Value
            End If
        Next FOCell
        NetWainscotTrimLength = b.bLength + NetWainscotTrimLength
    End If
    'subtract width of OH doors and Pdoors where there is Masonry Wainscot
    NetWainscotTrimLength = NetWainscotTrimLength - TempDoorWidth
    'convert
    NetWainscotTrimLength = NetWainscotTrimLength * 12
    'generate trim collection
    Call TrimPieceCalc(WainscotTrimPieces, NetWainscotTrimLength, "Masonry Wainscot")
    'Add qualities
    For Each TrimPiece In WainscotTrimPieces
        TrimPiece.tShape = "R-Loc"
        TrimPiece.Color = EstSht.Range("Wainscot_tColor").Value
    Next TrimPiece
End If
''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Gutters & Downspouts
''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Gutters
If b.Gutters = True Then
    Set GutterPieces = New Collection
    '''calculate gutter length
    'start with building length
    NetGutterLength = bLength * 12
    'add endwall overhangs or extensions
    NetGutterLength = NetGutterLength + b.e1Extension + b.e1ExtensionOverhang + b.e3Extension + b.e3ExtensionOverhang + b.e3Extension + b.e1Overhang + b.e3Overhang
    'if a gable roof, multiply by 2 to account for gutter along both sidewalls
    If rShape = "Gable" Then NetGutterLength = NetGutterLength * 2
    'generate gutter piece collection (done in the same way as trim, so using the trim piece sub)
    Call TrimPieceCalc(GutterPieces, NetGutterLength, "Gutter", PitchString)
    'add qualities
    For Each GutterPiece In GutterPieces
        GutterPiece.tShape = "R-Loc"
        GutterPiece.Color = GutterColor
        If GutterPiece.tLength = 244 Then
            GutterPiece.tLength = 21 * 12
            GutterPiece.tMeasurement = ImperialMeasurementFormat(GutterPiece.tLength)
        End If
    Next GutterPiece
    'end caps
    If rShape = "Gable" Then
        GutterEndCapQty = 4
    ElseIf rShape = "Single Slope" Then
        GutterEndCapQty = 2
    End If
    ''straps (same as the qty of bottom roof sheets)
    'first, building length in ft
    GutterStrapQty = ((bLength * 12) + b.e1Extension + b.e1ExtensionOverhang + b.e3Extension + b.e3ExtensionOverhang + b.e3Extension + b.e1Overhang + b.e3Overhang) / 12
    'divide by 3', round up
    GutterStrapQty = Application.WorksheetFunction.RoundUp(GutterStrapQty / 3, 0)
    'multiply by 2 if a gable roof
    If rShape = "Gable" Then GutterStrapQty = GutterStrapQty * 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Downspouts
    Set DownspoutPieces = New Collection
    'find downspout quantity
    If rShape = "Gable" Then
        DownspoutQty = (BayCount + 1) * 2
    ElseIf rShape = "Single Slope" Then
        DownspoutQty = BayCount + 1
    End If
    'find first kickout piece height
    Set DownspoutPiece = New clsTrim
    DownspoutPiece.tType = "Square Downspout W/ Kickout"    'set type as downspout with kickout
    Select Case bHeight * 12
    Case Is <= 122
        DownspoutPiece.tMeasurement = "10'2"""
        DownspoutPiece.tLength = 122
    Case Is <= 146
        DownspoutPiece.tMeasurement = "12'2"""
        DownspoutPiece.tLength = 146
    Case Is <= 170
        DownspoutPiece.tMeasurement = "14'2"""
        DownspoutPiece.tLength = 170
    Case Is <= 194
        DownspoutPiece.tMeasurement = "16'2"""
        DownspoutPiece.tLength = 194
    Case Is <= 218
        DownspoutPiece.tMeasurement = "18'2"""
        DownspoutPiece.tLength = 218
    Case Is <= 244
        DownspoutPiece.tMeasurement = "20'4"""
        DownspoutPiece.tLength = 244
    Case Else       'greater than 20'4
        RemainingHeight = (bHeight * 12) - 242
        DownspoutPiece.tMeasurement = "20'4"""
        DownspoutPiece.tLength = 244
    End Select
    'set quantity, shape, color
    DownspoutPiece.Quantity = DownspoutQty
    DownspoutPiece.tShape = "R-Loc"
    DownspoutPiece.Color = DownspoutColor
    DownspoutPieces.Add DownspoutPiece
    'find the rest of pieces
    If RemainingHeight <> 0 Then Call TrimPieceCalc(DownspoutPieces, RemainingHeight, "Downspout", , DownspoutQty)
    'update downspout without kickout shape
    For Each DownspoutPiece In DownspoutPieces
        If DownspoutPiece.tType = "Square Downspout W/O Kickout" Then
            DownspoutPiece.tShape = "R-Loc"
            DownspoutPiece.Color = DownspoutColor
        End If
    Next DownspoutPiece
    '''' Straps
    ' find straps per downspout
    'reset building height. first strap as at 12'
    RemainingHeight = bHeight - 12
    DownspoutStrapQty = 1
    Do While RemainingHeight > 0
        'strap every 7'
        RemainingHeight = RemainingHeight - 7
        DownspoutStrapQty = DownspoutStrapQty + 1
    Loop
    DownspoutStrapQty = DownspoutStrapQty * DownspoutQty
    '''' Pop Rivits
    '# of rivits = (gutter piece qty *2*10) rounded up to 100
    For Each GutterPiece In GutterPieces
        PopRivitQty = PopRivitQty + GutterPiece.Quantity
    Next GutterPiece
    PopRivitQty = PopRivitQty * 2 * 10
    'round up to nearest 100
    PopRivitQty = Application.WorksheetFunction.RoundUp(PopRivitQty / 100, 0) * 100
End If
''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Translucent Wall Panels & Skylights
''''
''' skylight qty
'1 per skylight
If EstSht.Range("SkylightQty").Value > 0 Then SkylightPanelQty = EstSht.Range("SkylightQty").Value
''translucent wall panel qty
'half per translucent panel
SkylightPanelQty = SkylightPanelQty + Application.WorksheetFunction.RoundUp(EstSht.Range("TranslucentWallPanelQty").Value / 2, 0)



''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Extensions
''''
Set e1ExtensionPanels = New Collection
Set s2ExtensionPanels = New Collection
Set e3ExtensionPanels = New Collection
Set s4ExtensionPanels = New Collection

Set e1SoffitPanels = New Collection
Set e1ExtensionSoffitTrim = New Collection
Set s2SoffitPanels = New Collection
Set s2ExtensionSoffitTrim = New Collection
Set e3SoffitPanels = New Collection
Set e3ExtensionSoffitTrim = New Collection
Set s4SoffitPanels = New Collection
Set s4ExtensionSoffitTrim = New Collection

'''' e1 Gable Extension
If e1GableExtensionSection = True Then Call ExtensionPanelGen(e1ExtensionPanels, b, "e1_GableExtension", s2RoofPanels, s4RoofPanels)

'2x8 inside angle
If e1GableExtensionSoffit = True Then
    'add soffit panels to extension panel collection
    Call SoffitGen(e1SoffitPanels, e1ExtensionSoffitTrim, "e1_GableExtension", b, s2RoofPanels, s4RoofPanels)
    For Each SoffitPanel In e1SoffitPanels
        'add to extension panels
        e1ExtensionPanels.Add SoffitPanel
    Next SoffitPanel
    'consolodate duplicate panels
    Call DuplicateMaterialRemoval(e1ExtensionPanels, "Panel")
'    'correct trim color, shape
    For Each SoffitTrim In e1ExtensionSoffitTrim
        If SoffitTrim.tType <> "2x6 Outside Angle Trim" Then SoffitTrim.tShape = EstSht.Range("e1_GableExtensionSoffit").offset(0, 1).Value
        SoffitTrim.Color = EstSht.Range("e1_GableExtensionSoffit").offset(0, 4).Value
    Next SoffitTrim
End If


'''' e3 Gable Extension
If e3GableExtensionSection = True Then Call ExtensionPanelGen(e3ExtensionPanels, b, "e3_GableExtension", s2RoofPanels, s4RoofPanels)
If e3GableExtensionSoffit = True Then
    Call SoffitGen(e3SoffitPanels, e3ExtensionSoffitTrim, "e3_GableExtension", b, s2RoofPanels, s4RoofPanels)
    'add soffit panels to extension panel collection
    For Each SoffitPanel In e3SoffitPanels
        'add to extension panels
        e3ExtensionPanels.Add SoffitPanel
    Next SoffitPanel
    'consolodate duplicate panels
    Call DuplicateMaterialRemoval(e3ExtensionPanels, "Panel")
'    'correct trim color, shape
    For Each SoffitTrim In e3ExtensionSoffitTrim
        If SoffitTrim.tType <> "2x6 Outside Angle Trim" Then SoffitTrim.tShape = EstSht.Range("e3_GableExtensionSoffit").offset(0, 1).Value
        SoffitTrim.Color = EstSht.Range("e3_GableExtensionSoffit").offset(0, 4).Value
    Next SoffitTrim
End If

's2 eave Extension
If s2EaveExtensionSection = True Then Call ExtensionPanelGen(s2ExtensionPanels, b, "s2_EaveExtension")
If s2EaveExtensionSoffit = True Then
    Call SoffitGen(s2SoffitPanels, s2ExtensionSoffitTrim, "s2_EaveExtension", b)
    'add soffit panels to extension panel collection
    For Each SoffitPanel In s2SoffitPanels
        'add to extension panels
        s2ExtensionPanels.Add SoffitPanel
    Next SoffitPanel
    'consolodate duplicate panels
    Call DuplicateMaterialRemoval(s2ExtensionPanels, "Panel")
    'correct trim color, shape
    For Each SoffitTrim In s2ExtensionSoffitTrim
        If SoffitTrim.tType <> "2x6 Outside Angle Trim" Then SoffitTrim.tShape = EstSht.Range("s2_EaveExtensionSoffit").offset(0, 1).Value
        SoffitTrim.Color = EstSht.Range("s2_EaveExtensionSoffit").offset(0, 4).Value
    Next SoffitTrim
End If

's4 eave Extension
If s4EaveExtensionSection = True Then Call ExtensionPanelGen(s4ExtensionPanels, b, "s4_EaveExtension")
If s4EaveExtensionSoffit = True Then
    Call SoffitGen(s4SoffitPanels, s4ExtensionSoffitTrim, "s4_EaveExtension", b)
    'add soffit panels to extension panel collection
    For Each SoffitPanel In s4SoffitPanels
        'add to extension panels
        s4ExtensionPanels.Add SoffitPanel
    Next SoffitPanel
    'consolodate duplicate panels
    Call DuplicateMaterialRemoval(s4ExtensionPanels, "Panel")
    'correct trim color, shape
    For Each SoffitTrim In s4ExtensionSoffitTrim
        If SoffitTrim.tType <> "2x6 Outside Angle Trim" Then SoffitTrim.tShape = EstSht.Range("s4_EaveExtensionSoffit").offset(0, 1).Value
        SoffitTrim.Color = EstSht.Range("s4_EaveExtensionSoffit").offset(0, 4).Value
    Next SoffitTrim
End If



''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Overhang Soffits
''''

'reset panel collections
Set e1SoffitPanels = New Collection
Set s2SoffitPanels = New Collection
Set e3SoffitPanels = New Collection
Set s4SoffitPanels = New Collection
' init overhang soffit trim collections
Set e1SoffitTrim = New Collection
Set s2SoffitTrim = New Collection
Set e3SoffitTrim = New Collection
Set s4SoffitTrim = New Collection

'e1 Gable overhang
If e1GableOverhangSoffit = True Then
    Call SoffitGen(e1SoffitPanels, e1SoffitTrim, "e1_GableOverhang", b, s2RoofPanels, s4RoofPanels)
    'update soffit trim color, shape
    For Each SoffitTrim In e1SoffitTrim
        If SoffitTrim.tType <> "2x6 Outside Angle Trim" Then SoffitTrim.tShape = EstSht.Range("e1_GableOverhangSoffit").offset(0, 1).Value
        SoffitTrim.Color = EstSht.Range("e1_GableOverhangSoffit").offset(0, 4).Value
    Next SoffitTrim
End If
'e3 Gable overhang
If e3GableOverhangSoffit = True Then
    Call SoffitGen(e3SoffitPanels, e3SoffitTrim, "e3_GableOverhang", b, s2RoofPanels, s4RoofPanels)
    'update soffit trim color, shape
    For Each SoffitTrim In e3SoffitTrim
        If SoffitTrim.tType <> "2x6 Outside Angle Trim" Then SoffitTrim.tShape = EstSht.Range("e3_GableOverhangSoffit").offset(0, 1).Value
        SoffitTrim.Color = EstSht.Range("e3_GableOverhangSoffit").offset(0, 4).Value
    Next SoffitTrim
End If
's2 eave overhang
If s2EaveOverhangSoffit = True Then
    Call SoffitGen(s2SoffitPanels, s2SoffitTrim, "s2_EaveOverhang", b, s2RoofPanels, s4RoofPanels)
    'update soffit trim color, shape
    For Each SoffitTrim In s2SoffitTrim
        If SoffitTrim.tType <> "2x6 Outside Angle Trim" Then SoffitTrim.tShape = EstSht.Range("s2_EaveOverhangSoffit").offset(0, 1).Value
        SoffitTrim.Color = EstSht.Range("s2_EaveOverhangSoffit").offset(0, 4).Value
    Next SoffitTrim
End If
's4 eave overhang
If s4EaveOverhangSoffit = True Then
    Call SoffitGen(s4SoffitPanels, s4SoffitTrim, "s4_EaveOverhang", b, s2RoofPanels, s4RoofPanels)
    'update soffit trim color, shape
    For Each SoffitTrim In s4SoffitTrim
        If SoffitTrim.tType <> "2x6 Outside Angle Trim" Then SoffitTrim.tShape = EstSht.Range("s4_EaveOverhangSoffit").offset(0, 1).Value
        SoffitTrim.Color = EstSht.Range("s4_EaveOverhangSoffit").offset(0, 4).Value
    Next SoffitTrim
End If


'''''''''''''''''''' 2x8 Outside Angle trim for gable extensions without soffit
With b
'Generate 2x8 Inside angle trim
    If b.rShape = "Single Slope" Then
        If e1GableExtensionSection = True And e1GableExtensionSoffit = False Then
            Set e1InsideAngleTrim = New Collection
            Call TrimPieceCalc(e1InsideAngleTrim, .s2RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength + .s2ExtensionOverhangRafterLength + .s4ExtensionOverhangRafterLength, "Inside Angle", , , b)
        End If
        If e3GableExtensionSection = True And e3GableExtensionSoffit = False Then
            Set e3InsideAngleTrim = New Collection
            Call TrimPieceCalc(e3InsideAngleTrim, .s2RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength + .s2ExtensionOverhangRafterLength + .s4ExtensionOverhangRafterLength, "Inside Angle", , , b)
        End If
    ElseIf .rShape = "Gable" Then
        If e1GableExtensionSection = True And e1GableExtensionSoffit = False Then
            Set e1InsideAngleTrim = New Collection
            Call TrimPieceCalc(e1InsideAngleTrim, .s2RafterSheetLength + .s4RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength + .s2ExtensionOverhangRafterLength + .s4ExtensionOverhangRafterLength, "Inside Angle", , , b)
        End If
        If e3GableExtensionSection = True And e3GableExtensionSoffit = False Then
            Set e3InsideAngleTrim = New Collection
             Call TrimPieceCalc(e3InsideAngleTrim, .s2RafterSheetLength + .s4RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength + .s2ExtensionOverhangRafterLength + .s4ExtensionOverhangRafterLength, "Inside Angle", , , b)
        End If
    End If
    If Not e1InsideAngleTrim Is Nothing Then
        For Each TrimPiece In e1InsideAngleTrim
            TrimPiece.Color = .rPanelColor
        Next TrimPiece
    End If
    If Not e3InsideAngleTrim Is Nothing Then
        For Each TrimPiece In e3InsideAngleTrim
            TrimPiece.Color = .rPanelColor
        Next TrimPiece
    End If
End With

''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Fasteners
''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Roof Screws
'calculate panel overlaps
If rShape = "Single Slope" Then
    rOverlaps = (s2RoofPanels.Count - 1)
ElseIf rShape = "Gable" Then
    rOverlaps = (s2RoofPanels.Count - 1) + (s4RoofPanels.Count - 1)
End If
'''extension Overlaps
'rOverlaps = rOverlaps + (s2ExtensionPanels.Count - 1) + (s4ExtensionPanels.Count - 1)
'generate roof screws
Call RoofScrewGen(rTekScrewQty, rLapScrewQty, b, rOverlaps)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sidewall Screws
'sidewall panel overlaps
If s2SidewallPanels.Count > 0 Then sOverlaps = s2SidewallPanels.Count - 1
If s4SidewallPanels.Count > 0 Then sOverlaps = sOverlaps + s4SidewallPanels.Count - 1
'endwall overlaps
eOverlaps = b.e1WallPanelOverlaps + b.e3WallPanelOverlaps
'generate screws
Call WallScrewGen(wTekScrewQty, wLapScrewQty, b, sOverlaps, eOverlaps)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Trim Screws
Set TrimScrews = New Collection
Call TrimScrewCalc(TrimScrews, RakeTrimPieces, b)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Soffit Screws
If e1GableOverhangSoffit = True Then Call SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "e1_GableOverhang", b)
If s2EaveOverhangSoffit = True Then Call SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "s2_EaveOverhang", b)
If e3GableOverhangSoffit = True Then Call SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "e3_GableOverhang", b)
If s4EaveOverhangSoffit = True Then Call SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "s4_EaveOverhang", b)
If e1GableExtensionSoffit = True Then Call SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "e1_GableExtension", b)
If s2EaveExtensionSoffit = True Then Call SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "s2_EaveExtension", b)
If e3GableExtensionSoffit = True Then Call SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "e3_GableExtension", b)
If s4EaveExtensionSoffit = True Then Call SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "s4_EaveExtension", b)
'round up to nearest 250
SoffitScrewQty = Application.WorksheetFunction.RoundUp(SoffitScrewQty / 250, 0) * 250

''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Miscellaneous
''''

'Butyl tape, inside closures, and outside closures
Call MiscMaterialCalc(ButylTapeQty, InsideClosureQty, OutsideClosureQty, b, rOverlaps)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''     MATERIALS LIST OUTPUT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'delete old output sheet
Application.DisplayAlerts = False
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Employee Materials List" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n
Application.DisplayAlerts = True

'set new output sheet
MatShtTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set MatSht = ThisWorkbook.Sheets("MaterialsListTmp (2)")
'rename
MatSht.Name = "Employee Materials List"
MatSht.Visible = xlSheetVisible

'combined material collections
Set PanelCollection = New Collection
Set TrimCollection = New Collection
Set MiscCollection = New Collection

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' OUTPUT '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


With MatSht
    '''''''''''''''''''''''''''''''''''''' Roof Panels ''''''''''''''''''''''''''''''''''
    
    ''''''' Sidewall 2 Roof Panels ''''''''
    Call MatListSectionWrite(MatSht, .Range("s2_RoofSheetQtyCell1"), s2RoofPanels, "Panel")
    
    ''''''' Sidewall 4  Roof Panels ''''''''
    'delete if a single slope
    If rShape = "Single Slope" Then
        .Range("s4_RoofSheetQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    'output for gable roof
    ElseIf rShape = "Gable" Then
        Call MatListSectionWrite(MatSht, .Range("s4_RoofSheetQtyCell1"), s4RoofPanels, "Panel")
    End If
    
    ''''''''''''''''''''''''''''''''''''''''' ridge caps
    'delete if a single slope
    If rShape = "Single Slope" Then
        .Range("Roof_RidgeCapQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    'output for gable roof
    ElseIf rShape = "Gable" Then
        Set WriteCell = .Range("Roof_RidgeCapQtyCell1")
        WriteCell.Value = RidgeCapQty
        WriteCell.offset(0, 1).Value = "Formed Ridge Cap " & PitchString
        WriteCell.offset(0, 3).Value = "3'"
        WriteCell.offset(0, 4).Value = rColor
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''' sidewall panels
    '' sidewall 2
    Call MatListSectionWrite(MatSht, .Range("s2_SidewallSheetQtyCell1"), s2SidewallPanels, "Panel")
    
    '' sidewall 4
    Call MatListSectionWrite(MatSht, .Range("s4_SidewallSheetQtyCell1"), s4SidewallPanels, "Panel")
        
    
    ''''''''''''''''''''''''''''''''''''' endwall panels
    ''' endwall #1
    Call MatListSectionWrite(MatSht, .Range("e1_EndwallSheetQtyCell1"), e1EndwallPanels, "Panel")

    ''' endwall #3
    Call MatListSectionWrite(MatSht, .Range("e3_EndwallSheetQtyCell1"), e3EndwallPanels, "Panel")

    ''''''''''''''''''''''''''''''''''''''''' Liner Panels
    If LinerPanelsSection = False Then
        Range(.Range("e1_LinerPanelsQtyCell1").offset(-5, 0), .Range("Roof_LinerPanelsQtyCell1").offset(1, 0)).EntireRow.Delete
    Else
        '''''''''''write liner panels
        Set WriteCell = .Range("e1_LinerPanelsQtyCell1")
        Call MatListSectionWrite(MatSht, .Range("e1_LinerPanelsQtyCell1"), e1LinerPanels, "Panel")
        Call MatListSectionWrite(MatSht, .Range("e3_LinerPanelsQtyCell1"), e3LinerPanels, "Panel")
        Call MatListSectionWrite(MatSht, .Range("s2_LinerPanelsQtyCell1"), s2LinerPanels, "Panel")
        Call MatListSectionWrite(MatSht, .Range("s4_LinerPanelsQtyCell1"), s4LinerPanels, "Panel")
        Call MatListSectionWrite(MatSht, .Range("Roof_LinerPanelsQtyCell1"), RoofLinerPanels, "Panel")
        'clean up unused sections
        If .Range("e1_LinerPanelsQtyCell1").Value = "" Then .Range("e1_LinerPanelsQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
        If .Range("e3_LinerPanelsQtyCell1").Value = "" Then .Range("e3_LinerPanelsQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
        If .Range("s2_LinerPanelsQtyCell1").Value = "" Then .Range("s2_LinerPanelsQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
        If .Range("s4_LinerPanelsQtyCell1").Value = "" Then .Range("s4_LinerPanelsQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
        If .Range("Roof_LinerPanelsQtyCell1").Value = "" Then .Range("Roof_LinerPanelsQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''' Trim
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Rake
    Call MatListSectionWrite(MatSht, .Range("RakeTrimQtyCell1"), RakeTrimPieces, "Trim")
    ''' Eave
    Call MatListSectionWrite(MatSht, .Range("EaveTrimQtyCell1"), EaveTrimPieces, "Trim")
    ''' Outside Corner
    Call MatListSectionWrite(MatSht, .Range("OutsideCornerTrimQtyCell1"), OutsideCornerTrimPieces, "Trim")
    ''' Base Trim
    If b.BaseTrim = False Then
        .Range("BaseTrimQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Call MatListSectionWrite(MatSht, .Range("BaseTrimQtyCell1"), BaseTrimPieces, "Trim")
    End If
    ''' Wainscot Trim
    If EstSht.Range("Wainscot").Value = "No" Then
        .Range("WainscotTrimQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Call MatListSectionWrite(MatSht, .Range("WainscotTrimQtyCell1"), WainscotTrimPieces, "Trim")
    End If
    ''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FO Trim
    ''''
    Call FOMaterialGen(MatSht, TrimCollection, MiscCollection)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''' Gutters and Downspouts
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' gutters
    'delete entier gutters section if no gutters
    If Gutters = False Then
        Range(.Range("GutterQtyCell1").offset(-5, 0), .Range("DownspoutQtyCell1").offset(2, 0)).EntireRow.Delete
    Else
        Set WriteCell = .Range("GutterQtyCell1")
        '''gutter pieces
        For Each GutterPiece In GutterPieces
            'insert new row if not the first write cell in the section
            If WriteCell <> .Range("GutterQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            'add piece
            WriteCell.Value = GutterPiece.Quantity
            WriteCell.offset(0, 1).Value = GutterPiece.tShape
            WriteCell.offset(0, 2).Value = GutterPiece.tType
            WriteCell.offset(0, 3).Value = GutterPiece.tMeasurement
            WriteCell.offset(0, 4).Value = GutterPiece.Color
            'update write cell
            Set WriteCell = WriteCell.offset(1, 0)
        Next GutterPiece
        '''end caps
        .Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = GutterEndCapQty
        WriteCell.offset(0, 1).Value = "R-Loc"
        WriteCell.offset(0, 2).Value = "Sculptured Gutter End Cap"
        WriteCell.offset(0, 3).Value = "N/A"
        WriteCell.offset(0, 4).Value = GutterColor
        Set WriteCell = WriteCell.offset(1, 0)
        '''straps
        .Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = GutterStrapQty
        WriteCell.offset(0, 1).Value = "R-Loc"
        WriteCell.offset(0, 2).Value = "Gutter Strap 9"""
        WriteCell.offset(0, 3).Value = "N/A"
        WriteCell.offset(0, 4).Value = GutterColor
        Set WriteCell = WriteCell.offset(1, 0)
          
        '''' Downspouts
        Set WriteCell = .Range("DownspoutQtyCell1")
        '''Downspout pieces
        For Each DownspoutPiece In DownspoutPieces
            'insert new row if not the first write cell in the section
            If WriteCell <> .Range("DownspoutQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            'add piece
            WriteCell.Value = DownspoutPiece.Quantity
            WriteCell.offset(0, 1).Value = DownspoutPiece.tShape
            WriteCell.offset(0, 2).Value = DownspoutPiece.tType
            WriteCell.offset(0, 3).Value = DownspoutPiece.tMeasurement
            WriteCell.offset(0, 4).Value = DownspoutPiece.Color
            'update write cell
            Set WriteCell = WriteCell.offset(1, 0)
        Next DownspoutPiece
        ''' downspout straps
        .Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = DownspoutStrapQty
        WriteCell.offset(0, 1).Value = "N/A"
        WriteCell.offset(0, 2).Value = "Downspout Strap"
        WriteCell.offset(0, 3).Value = "N/A"
        WriteCell.offset(0, 4).Value = DownspoutColor
        Set WriteCell = WriteCell.offset(1, 0)
        ''' pop rivits
        .Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = PopRivitQty
        WriteCell.offset(0, 1).Value = "N/A"
        WriteCell.offset(0, 2).Value = "Pop Rivets"
        WriteCell.offset(0, 3).Value = "1"""
        WriteCell.offset(0, 4).Value = DownspoutColor
        Set WriteCell = WriteCell.offset(1, 0)
    End If
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''' Additional Options
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Skylights & Translucent Wall Panels
    If SkylightPanelQty = 0 Then
        'check if deleting entire section
        If e1GableOverhangSection = False And s2EaveOverhangSection = False And e3GableOverhangSection = False And _
        s4EaveOverhangSection = False And e1GableExtensionSection = False And s2EaveExtensionSection = False And _
        e3GableExtensionSection = False And s4EaveExtensionSection = False Then
            'delete "Additional Options" Section heading as well
            .Range("skylightPanelQtyCell1").offset(-4, 0).Resize(6, 1).EntireRow.Delete
        Else
            'just delete skylights and translucent wall panels heading
            .Range("SkylightPanelQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
        End If
    Else
        Set WriteCell = .Range("SkylightPanelQtyCell1")
        WriteCell.Value = SkylightPanelQty
        WriteCell.offset(0, 1).Value = "R-Loc"
        WriteCell.offset(0, 2).Value = "Skylights, Fiberglass, White"
        WriteCell.offset(0, 3).Value = "12'"
        WriteCell.offset(0, 4).Value = "N/A"
    End If
    'e1 Gable overhang Soffit
    If e1GableOverhangSection = False Then
        .Range("e1_GableOverhangMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        'Call MatListSectionWrite(MatSht, .Range("e1_GableOverhangMatQtyCell1"), e1SoffitPanels, "Panel")
        Set WriteCell = .Range("e1_GableOverhangMatQtyCell1")
        ' Soffit Panels
        For Each SoffitPanel In e1SoffitPanels
            If WriteCell <> .Range("e1_GableOverhangMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitPanel.Quantity
            WriteCell.offset(0, 1).Value = SoffitPanel.PanelShape
            WriteCell.offset(0, 2).Value = SoffitPanel.PanelType
            WriteCell.offset(0, 3).Value = SoffitPanel.PanelMeasurement
            WriteCell.offset(0, 4).Value = SoffitPanel.PanelColor
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitPanel
        ' Soffit Trim
        For Each SoffitTrim In e1SoffitTrim
            If WriteCell <> .Range("e1_GableOverhangMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitTrim.Quantity
            WriteCell.offset(0, 1).Value = SoffitTrim.tShape
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement
            WriteCell.offset(0, 4).Value = SoffitTrim.Color
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitTrim
    End If
    's2 eave overhang Soffit
    If s2EaveOverhangSection = False Then
        .Range("s2_EaveOverhangMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("s2_EaveOverhangMatQtyCell1")
        ' Soffit Panels
        For Each SoffitPanel In s2SoffitPanels
            If WriteCell <> .Range("s2_EaveOverhangMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitPanel.Quantity
            WriteCell.offset(0, 1).Value = SoffitPanel.PanelShape
            WriteCell.offset(0, 2).Value = SoffitPanel.PanelType
            WriteCell.offset(0, 3).Value = SoffitPanel.PanelMeasurement
            WriteCell.offset(0, 4).Value = SoffitPanel.PanelColor
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitPanel
        ' Soffit Trim
        For Each SoffitTrim In s2SoffitTrim
            If WriteCell <> .Range("s2_EaveOverhangMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitTrim.Quantity
            WriteCell.offset(0, 1).Value = SoffitTrim.tShape
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement
            WriteCell.offset(0, 4).Value = SoffitTrim.Color
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitTrim
    End If
    'e3 Gable overhang Soffit
    If e3GableOverhangSection = False Then
        .Range("e3_GableOverhangMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("e3_GableOverhangMatQtyCell1")
        ' Soffit Panels
        For Each SoffitPanel In e3SoffitPanels
            If WriteCell <> .Range("e3_GableOverhangMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitPanel.Quantity
            WriteCell.offset(0, 1).Value = SoffitPanel.PanelShape
            WriteCell.offset(0, 2).Value = SoffitPanel.PanelType
            WriteCell.offset(0, 3).Value = SoffitPanel.PanelMeasurement
            WriteCell.offset(0, 4).Value = SoffitPanel.PanelColor
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitPanel
        ' Soffit Trim
        For Each SoffitTrim In e3SoffitTrim
            If WriteCell <> .Range("e3_GableOverhangMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitTrim.Quantity
            WriteCell.offset(0, 1).Value = SoffitTrim.tShape
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement
            WriteCell.offset(0, 4).Value = SoffitTrim.Color
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitTrim
    End If
    's4 eave overhang Soffit
    If s4EaveOverhangSection = False Then
        .Range("s4_EaveOverhangMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("s4_EaveOverhangMatQtyCell1")
        ' Soffit Panels
        For Each SoffitPanel In s4SoffitPanels
            If WriteCell <> .Range("s4_EaveOverhangMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitPanel.Quantity
            WriteCell.offset(0, 1).Value = SoffitPanel.PanelShape
            WriteCell.offset(0, 2).Value = SoffitPanel.PanelType
            WriteCell.offset(0, 3).Value = SoffitPanel.PanelMeasurement
            WriteCell.offset(0, 4).Value = SoffitPanel.PanelColor
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitPanel
        ' Soffit Trim
        For Each SoffitTrim In s4SoffitTrim
            If WriteCell <> .Range("s4_EaveOverhangMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitTrim.Quantity
            WriteCell.offset(0, 1).Value = SoffitTrim.tShape
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement
            WriteCell.offset(0, 4).Value = SoffitTrim.Color
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitTrim
    End If
    'e1 Gable Extension
    If e1GableExtensionSection = False Then
        .Range("e1_GableExtensionMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("e1_GableExtensionMatQtyCell1")
        'Extension and Soffit Panels
        For Each ExtensionPanel In e1ExtensionPanels
            If WriteCell <> .Range("e1_GableExtensionMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = ExtensionPanel.Quantity
            WriteCell.offset(0, 1).Value = ExtensionPanel.PanelShape
            WriteCell.offset(0, 2).Value = ExtensionPanel.PanelType
            WriteCell.offset(0, 3).Value = ExtensionPanel.PanelMeasurement
            WriteCell.offset(0, 4).Value = ExtensionPanel.PanelColor
            Set WriteCell = WriteCell.offset(1, 0)
        Next ExtensionPanel
        '2x8 inside angle
        If Not e1InsideAngleTrim Is Nothing Then
            For Each TrimPiece In e1InsideAngleTrim
                If WriteCell <> .Range("e1_GableExtensionMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
                WriteCell.Value = TrimPiece.Quantity
                WriteCell.offset(0, 1).Value = TrimPiece.tShape
                WriteCell.offset(0, 2).Value = TrimPiece.tType
                WriteCell.offset(0, 3).Value = TrimPiece.tMeasurement
                WriteCell.offset(0, 4).Value = TrimPiece.Color
                Set WriteCell = WriteCell.offset(1, 0)
            Next TrimPiece
        End If
        ' Soffit Trim
        For Each SoffitTrim In e1ExtensionSoffitTrim
            If WriteCell <> .Range("e1_GableExtensionMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitTrim.Quantity
            WriteCell.offset(0, 1).Value = SoffitTrim.tShape
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement
            WriteCell.offset(0, 4).Value = SoffitTrim.Color
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitTrim
    End If
    's2 eave Extension
    If s2EaveExtensionSection = False Then
        .Range("s2_EaveExtensionMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("s2_EaveExtensionMatQtyCell1")
'       'Extension and Soffit Panels
        For Each ExtensionPanel In s2ExtensionPanels
            If WriteCell.Address <> .Range("s2_EaveExtensionMatQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = ExtensionPanel.Quantity
            WriteCell.offset(0, 1).Value = ExtensionPanel.PanelShape
            WriteCell.offset(0, 2).Value = ExtensionPanel.PanelType
            WriteCell.offset(0, 3).Value = ExtensionPanel.PanelMeasurement
            WriteCell.offset(0, 4).Value = ExtensionPanel.PanelColor
            Set WriteCell = WriteCell.offset(1, 0)
        Next ExtensionPanel
        ' Soffit Trim
        For Each SoffitTrim In s2ExtensionSoffitTrim
            If WriteCell.Address <> .Range("s2_EaveExtensionMatQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitTrim.Quantity
            WriteCell.offset(0, 1).Value = SoffitTrim.tShape
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement
            WriteCell.offset(0, 4).Value = SoffitTrim.Color
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitTrim
    End If
    'e3 Gable Extension
    If e3GableExtensionSection = False Then
        .Range("e3_GableExtensionMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("e3_GableExtensionMatQtyCell1")
        'Extension and Soffit Panels
        For Each ExtensionPanel In e3ExtensionPanels
            If WriteCell <> .Range("e3_GableExtensionMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = ExtensionPanel.Quantity
            WriteCell.offset(0, 1).Value = ExtensionPanel.PanelShape
            WriteCell.offset(0, 2).Value = ExtensionPanel.PanelType
            WriteCell.offset(0, 3).Value = ExtensionPanel.PanelMeasurement
            WriteCell.offset(0, 4).Value = ExtensionPanel.PanelColor
            Set WriteCell = WriteCell.offset(1, 0)
        Next ExtensionPanel
        '2x8 inside angle
        If Not e3InsideAngleTrim Is Nothing Then
            For Each TrimPiece In e3InsideAngleTrim
                If WriteCell <> .Range("e3_GableExtensionMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
                WriteCell.Value = TrimPiece.Quantity
                WriteCell.offset(0, 1).Value = TrimPiece.tShape
                WriteCell.offset(0, 2).Value = TrimPiece.tType
                WriteCell.offset(0, 3).Value = TrimPiece.tMeasurement
                WriteCell.offset(0, 4).Value = TrimPiece.Color
                Set WriteCell = WriteCell.offset(1, 0)
            Next TrimPiece
        End If
        ' Soffit Trim
        For Each SoffitTrim In e3ExtensionSoffitTrim
            If WriteCell <> .Range("e3_GableExtensionMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitTrim.Quantity
            WriteCell.offset(0, 1).Value = SoffitTrim.tShape
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement
            WriteCell.offset(0, 4).Value = SoffitTrim.Color
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitTrim
    End If
    's4 eave Extension
    If s4EaveExtensionSection = False Then
        .Range("s4_EaveExtensionMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("s4_EaveExtensionMatQtyCell1")
'       'Extension and Soffit Panels
        For Each ExtensionPanel In s4ExtensionPanels
            If WriteCell.Address <> .Range("s4_EaveExtensionMatQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = ExtensionPanel.Quantity
            WriteCell.offset(0, 1).Value = ExtensionPanel.PanelShape
            WriteCell.offset(0, 2).Value = ExtensionPanel.PanelType
            WriteCell.offset(0, 3).Value = ExtensionPanel.PanelMeasurement
            WriteCell.offset(0, 4).Value = ExtensionPanel.PanelColor
            Set WriteCell = WriteCell.offset(1, 0)
        Next ExtensionPanel
        ' Soffit Trim
        For Each SoffitTrim In s4ExtensionSoffitTrim
            If WriteCell.Address <> .Range("s4_EaveExtensionMatQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
            WriteCell.Value = SoffitTrim.Quantity
            WriteCell.offset(0, 1).Value = SoffitTrim.tShape
            WriteCell.offset(0, 2).Value = SoffitTrim.tType
            WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement
            WriteCell.offset(0, 4).Value = SoffitTrim.Color
            Set WriteCell = WriteCell.offset(1, 0)
        Next SoffitTrim
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Fasteners
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''' roof screws
    Set WriteCell = .Range("RoofScrewsQtyCell1")
    'tek screws
    WriteCell.Value = rTekScrewQty
    WriteCell.offset(0, 1).Value = "Tek Screws"
    WriteCell.offset(0, 3).Value = "1.25"""
    WriteCell.offset(0, 4).Value = rColor
    Set WriteCell = WriteCell.offset(1, 0)
    'lap screws
    .Rows(WriteCell.Row + 1).Insert
    WriteCell.Value = rLapScrewQty
    WriteCell.offset(0, 1).Value = "Lap Screws"
    WriteCell.offset(0, 3).Value = ".875"""
    WriteCell.offset(0, 4).Value = rColor
    ''''''''''''''''''''''''''''''''''''''''''''''''' wall screws
    Set WriteCell = .Range("WallScrewsQtyCell1")
    'tek screws
    WriteCell.Value = wTekScrewQty
    WriteCell.offset(0, 1).Value = "Tek Screws"
    WriteCell.offset(0, 3).Value = "1.25"""
    WriteCell.offset(0, 4).Value = wColor
    Set WriteCell = WriteCell.offset(1, 0)
    'lap screws
    .Rows(WriteCell.Row + 1).Insert
    WriteCell.Value = wLapScrewQty
    WriteCell.offset(0, 1).Value = "Lap Screws"
    WriteCell.offset(0, 3).Value = ".875"""
    WriteCell.offset(0, 4).Value = wColor
    ''''''''''''''''''''''''''''''''''''''''''''''''' Trim screws
    Set WriteCell = .Range("TrimScrewsQtyCell1")
    'tek screws
    For Each Screw In TrimScrews
        If WriteCell.Address <> .Range("TrimScrewsQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = Screw.Quantity
        WriteCell.offset(0, 1).Value = "Tek Screws"
        WriteCell.offset(0, 3).Value = "1.25"""
        WriteCell.offset(0, 4).Value = Screw.Color
        Set WriteCell = WriteCell.offset(1, 0)
    Next Screw
    'lap screws (duplicate colors/quantities)
    For Each Screw In TrimScrews
        If WriteCell.Address <> .Range("TrimScrewsQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = Screw.Quantity
        WriteCell.offset(0, 1).Value = "Lap Screws"
        WriteCell.offset(0, 3).Value = ".875"""
        WriteCell.offset(0, 4).Value = Screw.Color
        Set WriteCell = WriteCell.offset(1, 0)
    Next Screw

    ''''''''''''''''''''''''''''''''''''''''''''''''' Soffit Screws
    If SoffitScrewQty = 0 Then
        .Range("SoffitScrewsQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("SoffitScrewsQtyCell1")
        'tek screws
        WriteCell.Value = SoffitScrewQty
        WriteCell.offset(0, 1).Value = "Tek Screws"
        WriteCell.offset(0, 3).Value = "1.25"""
        WriteCell.offset(0, 4).Value = SoffitScrewColor
        Set WriteCell = WriteCell.offset(1, 0)
        'lap screws
        .Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = SoffitScrewQty
        WriteCell.offset(0, 1).Value = "Lap Screws"
        WriteCell.offset(0, 3).Value = ".875"""
        WriteCell.offset(0, 4).Value = SoffitScrewColor
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Miscellaneous
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Butyl Tape
    Set WriteCell = .Range("MiscMaterialsQtyCell1")
    WriteCell.Value = ButylTapeQty
    WriteCell.offset(0, 1).Value = "Butyl Tape"
    WriteCell.offset(0, 3).Value = "44'"
    WriteCell.offset(0, 4).Value = "N/A"
    Set WriteCell = WriteCell.offset(1, 0)
    'Inside Closures
    WriteCell.Value = InsideClosureQty
    WriteCell.offset(0, 1).Value = "Inside Closures"
    WriteCell.offset(0, 3).Value = "3'"
    WriteCell.offset(0, 4).Value = "N/A"
    Set WriteCell = WriteCell.offset(1, 0)
    'Outside Closures
    WriteCell.Value = OutsideClosureQty
    WriteCell.offset(0, 1).Value = "Outside Closures"
    WriteCell.offset(0, 3).Value = "3'"
    WriteCell.offset(0, 4).Value = "N/A"
    Set WriteCell = WriteCell.offset(1, 0)
End With

'autofit columns
MatSht.Columns.AutoFit


''''''''''''''''''''''''''''''''''''' Vendor Material List
'''''Roof Panels
For Each RoofPanel In s2RoofPanels
    PanelCollection.Add RoofPanel
Next RoofPanel
For Each RoofPanel In s4RoofPanels
    PanelCollection.Add RoofPanel
Next RoofPanel
' Ridge Caps
If RidgeCapQty <> 0 Then
    Set item = New clsMiscItem
    item.Quantity = RidgeCapQty
    item.Name = "Formed Ridge Cap " & PitchString
    item.Measurement = "3'"
    item.Color = rColor
    MiscCollection.Add item
End If
''''' Sidewall Panels
For Each Panel In s2SidewallPanels
    PanelCollection.Add Panel
Next Panel
For Each Panel In s4SidewallPanels
    PanelCollection.Add Panel
Next Panel
    
''''' Endwall panels
For Each Panel In e1EndwallPanels
    PanelCollection.Add Panel
Next Panel
For Each Panel In e3EndwallPanels
    PanelCollection.Add Panel
Next Panel
'liner panels
For Each Panel In e1LinerPanels
    PanelCollection.Add Panel
Next Panel
For Each Panel In e3LinerPanels
    PanelCollection.Add Panel
Next Panel
For Each Panel In s2LinerPanels
    PanelCollection.Add Panel
Next Panel
For Each Panel In s4LinerPanels
    PanelCollection.Add Panel
Next Panel
For Each Panel In RoofLinerPanels
    PanelCollection.Add Panel
Next Panel
''''' Trim
For Each TrimPiece In RakeTrimPieces
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In EaveTrimPieces
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In OutsideCornerTrimPieces
    TrimCollection.Add TrimPiece
Next TrimPiece
If b.BaseTrim = True Then
    For Each TrimPiece In BaseTrimPieces
        TrimCollection.Add TrimPiece
    Next TrimPiece
End If
''''' Gutters
If b.Gutters = True Then
    For Each GutterPiece In GutterPieces
        'trim item used for gutters
        TrimCollection.Add GutterPiece
    Next GutterPiece
    ' Gutter End Caps
    If GutterEndCapQty <> 0 Then
        Set item = New clsMiscItem
        item.Quantity = GutterEndCapQty
        item.Shape = "R-Loc"
        item.Name = "Sculptured Gutter End Cap"
        item.Measurement = "N/A"
        item.Color = GutterColor
        MiscCollection.Add item
    End If
    ' Gutter Straps
    If GutterStrapQty <> 0 Then
        Set item = New clsMiscItem
        item.Quantity = GutterStrapQty
        item.Shape = "R-Loc"
        item.Name = "Gutter Strap"
        item.Measurement = "9"""
        item.Color = GutterColor
        MiscCollection.Add item
    End If
    For Each DownspoutPiece In DownspoutPieces
        'trim item used for downspout pieces
        TrimCollection.Add DownspoutPiece
    Next DownspoutPiece
    ' Downspout Straps
    If DownspoutStrapQty <> 0 Then
        Set item = New clsMiscItem
        item.Quantity = DownspoutStrapQty
        item.Shape = "N/A"
        item.Name = "Downspout Strap"
        item.Measurement = "N/A"
        item.Color = DownspoutColor
        MiscCollection.Add item
    End If
    ' Pop Rivets
    If PopRivitQty <> 0 Then
        Set item = New clsMiscItem
        item.Quantity = PopRivitQty
        item.Shape = "N/A"
        item.Name = "Pop Rivets"
        item.Measurement = "1"""
        item.Color = DownspoutColor
        MiscCollection.Add item
    End If
End If
''''' Skylights and Translucent Wall Panels
If SkylightPanelQty <> 0 Then
    Set SkylightPanel = New clsPanel
    SkylightPanel.PanelShape = "R-Loc"
    SkylightPanel.PanelType = "Skylights, Fiberglass, White"
    SkylightPanel.PanelMeasurement = "12'"
    SkylightPanel.PanelLength = 12 * 12
    SkylightPanel.PanelColor = "N/A"
    SkylightPanel.Quantity = SkylightPanelQty
    PanelCollection.Add SkylightPanel
End If
''''' Overhangs, Extensions, and Soffits
'extension panels
For Each ExtensionPanel In e1ExtensionPanels
    PanelCollection.Add ExtensionPanel
Next ExtensionPanel
For Each ExtensionPanel In s2ExtensionPanels
    PanelCollection.Add ExtensionPanel
Next ExtensionPanel
For Each ExtensionPanel In e3ExtensionPanels
    PanelCollection.Add ExtensionPanel
Next ExtensionPanel
For Each ExtensionPanel In s4ExtensionPanels
    PanelCollection.Add ExtensionPanel
Next ExtensionPanel
'soffit panels
For Each SoffitPanel In e1SoffitPanels
    PanelCollection.Add SoffitPanel
Next SoffitPanel
For Each SoffitPanel In s2SoffitPanels
    PanelCollection.Add SoffitPanel
Next SoffitPanel
For Each SoffitPanel In e3SoffitPanels
    PanelCollection.Add SoffitPanel
Next SoffitPanel
For Each SoffitPanel In s4SoffitPanels
    PanelCollection.Add SoffitPanel
Next SoffitPanel
'soffit trim
For Each TrimPiece In e1SoffitTrim
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In s2SoffitTrim
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In e3SoffitTrim
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In s4SoffitTrim
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In e1ExtensionSoffitTrim
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In s2ExtensionSoffitTrim
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In e3ExtensionSoffitTrim
    TrimCollection.Add TrimPiece
Next TrimPiece
For Each TrimPiece In s4ExtensionSoffitTrim
    TrimCollection.Add TrimPiece
Next TrimPiece
''''' Fasteners and Miscelaneous
' Roof Screws
Set item = New clsMiscItem
item.Quantity = rTekScrewQty
item.Name = "Tek Screws"
item.Measurement = "1.25"""
item.Color = rColor
MiscCollection.Add item
Set item = New clsMiscItem
item.Quantity = rLapScrewQty
item.Name = "Lap Screws"
item.Measurement = ".875"""
item.Color = rColor
MiscCollection.Add item
' Wall Screws
Set item = New clsMiscItem
item.Quantity = wTekScrewQty
item.Name = "Tek Screws"
item.Measurement = "1.25"""
item.Color = wColor
MiscCollection.Add item
Set item = New clsMiscItem
item.Quantity = wLapScrewQty
item.Name = "Lap Screws"
item.Measurement = ".875"""
item.Color = wColor
MiscCollection.Add item
''' Trim Screws
'Tek
For Each Screw In TrimScrews
    Set item = New clsMiscItem
    item.Quantity = Screw.Quantity
    item.Name = "Tek Screws"
    item.Measurement = "1.25"""
    item.Color = Screw.Color
    MiscCollection.Add item
Next Screw
'Duplicate Lap
For Each Screw In TrimScrews
    Set item = New clsMiscItem
    item.Quantity = Screw.Quantity
    item.Name = "Lap Screws"
    item.Measurement = ".875"""
    item.Color = Screw.Color
    MiscCollection.Add item
Next Screw
''' Soffit Screws
If SoffitScrewQty <> 0 Then
    'Tek
    Set item = New clsMiscItem
    item.Quantity = SoffitScrewQty
    item.Name = "Tek Screws"
    item.Measurement = "1.25"""
    item.Color = SoffitScrewColor
    MiscCollection.Add item
    'Duplicate Lap
    Set item = New clsMiscItem
    item.Quantity = SoffitScrewQty
    item.Name = "Lap Screws"
    item.Measurement = ".875"""
    item.Color = SoffitScrewColor
    MiscCollection.Add item
End If
''''' Miscellaneous
'Butyl Tape
Set item = New clsMiscItem
item.Quantity = ButylTapeQty
item.Name = "Butyl Tape"
item.Measurement = "44'"
MiscCollection.Add item
'Inside Closures
Set item = New clsMiscItem
item.Quantity = InsideClosureQty
item.Name = "Inside Closures"
item.Measurement = "3'"
MiscCollection.Add item
'Outside Closures
Set item = New clsMiscItem
item.Quantity = OutsideClosureQty
item.Name = "Outside Closures"
item.Measurement = "3'"
MiscCollection.Add item


'write collections to building class
Set b.PanelCollection = PanelCollection
Set b.TrimCollection = TrimCollection
Set b.MiscMaterialsCollection = MiscCollection


' generate the rest of the misc. materials
Call MiscMaterialsGen.MiscMaterialCalc(b.MiscMaterialsCollection, WriteCell, b)


'''''generate vendor material list
'remove duplicates
Call DuplicateMaterialRemoval(b.PanelCollection, "Panel")
Call DuplicateMaterialRemoval(b.TrimCollection, "Trim")
Call DuplicateMaterialRemoval(b.MiscMaterialsCollection, "Misc")



Exit Sub

MissingRoofData:
MsgBox "Key information to determine the roofing materials is missing! Please check the template for missing data and try again.", vbExclamation, "Missing Data"
End

LargePanelDivision:
MsgBox "It has been calculated that more than 5 seperate panels will be needed to cover the rafter length of the roof. Please perform this calculation manually.", vbExclamation, "Program Design Exceeded"
End

TempDisabled:
MsgBox "It has been calculated that 4+ seperate panels will be needed to cover the rafter length of the roof. This calculation is currently disabled.", vbInformation, "Currently Disabled"
End

SingleSlopeDisabled:
MsgBox "Single slope calculations are currently disabled.", vbInformation, "Feature Disabled"
End

End Sub


Function ImperialMeasurementFormat(TotalInches As Double) As String
Dim Feet As Single
Dim Inches As Double
Dim InchFraction As Double
Dim InchFractString

Feet = Application.WorksheetFunction.RoundDown(TotalInches / 12, 0)
Inches = Application.WorksheetFunction.RoundDown((XLMod((TotalInches / 12), 1) * 12), 0)
InchFraction = Application.WorksheetFunction.MRound(XLMod((XLMod((TotalInches / 12), 1) * 12), 1), 1 / 16)

'add to inches if inch fraction = 1
If InchFraction = 1 Then
    Inches = Inches + 1
    InchFraction = 0
    'check if 12 inches
    If Inches = 12 Then
        Feet = Feet + 1
        Inches = 0
    End If
End If

''write values to formatting cell
With HiddenSht
    .Range("Inch_Fraction_Format").Value = InchFraction
    .Range("Inch_Format").Value = Inches
    .Range("Feet_Format").Value = Feet

    'write string
    If Inches = 0 And InchFraction = 0 Then
        ImperialMeasurementFormat = .Range("Feet_Format").Text & "'"
    ElseIf InchFraction = 0 Then
        'ImperialMeasurementFormat = .Range("Feet_Format").Text & "'" & " " & .Range("Inch_Format").Text & "''"
        ImperialMeasurementFormat = .Range("Feet_Format").Text & "'" & " " & .Range("Inch_Format").Text & "''"
    ElseIf Inches = 0 Then
        ImperialMeasurementFormat = .Range("Feet_Format").Text & "'" & " " & Trim(.Range("Inch_Fraction_Format").Text) & "''"
    Else
        ImperialMeasurementFormat = .Range("Feet_Format").Text & "'" & " " & .Range("Inch_Format").Text & " " & Trim(.Range("Inch_Fraction_Format").Text) & "''"
    End If
End With
    
        

End Function

Function XLMod(a, b)
    ' This replicates the Excel MOD function
    XLMod = a - b * Int(a / b)
End Function


Function ClosestWallPurlin(Height As Variant, Optional Direction As Integer, Optional NonstandardFloorPurlin As Boolean) As Double
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function returns the nearest value to a target
'INPUT: Pass the function a range of cells, a target value that you want to find a number closest to
' and an optional direction variable described below.
'OPTIONS: Set the optional variable Direction equal to 0 or blank to find the closest value
' Set equal to -1 to find the closest value below your target
' set equal to 1 to find the closest value above your target
'OUTPUT: The output is the number in the range closest to your target value.
' Because the output is a variant, the address of the closest number can also be returned when
' calling this function from another VBA macro.
Dim t As Variant
Dim u As Variant
Dim Purlins() As Variant
Dim Purlin As Variant
Dim pAbove As Integer
Dim pBelow As Integer
Purlins = Array(87.5, 147.5, 207.5, 267.5, 327.5, 387.5, 447.5, 507.5, 567.5, 627.5, 687.5, 747.5, 807.5, 867.5, 927.5, 987.5, 1047.5, 1107.5, 1167.5)

'Normal (Floor to Eave)
If NonstandardFloorPurlin = False Then
    t = 1.79769313486231E+308 'initialize
    'ClosestWallPurlin = "No value found"
    For Each Purlin In Purlins
        If IsNumeric(Purlin) Then
            u = Abs(Purlin - Height)
            If Direction > 0 And Purlin >= Height Then
                'only report if closer number is greater than the target
                If u < t Then
                    t = u
                    ClosestWallPurlin = Purlin
                End If
            ElseIf Direction < 0 And Purlin <= Height Then
                'only report if closer number is less than the target
                If u < t Then
                    t = u
                    ClosestWallPurlin = Purlin
                End If
            ElseIf Direction = 0 Then
                If u < t Then
                    t = u
                    ClosestWallPurlin = Purlin
                End If
            End If
        End If
    Next Purlin

'starting at bottom of partial wall instead of ground
ElseIf NonstandardFloorPurlin = True Then
    pBelow = Application.WorksheetFunction.RoundDown(Height / 5, 0) * 5
    pAbove = Application.WorksheetFunction.RoundUp(Height / 5, 0) * 5
    Select Case Direction
    Case -1
        ClosestWallPurlin = pBelow
    Case 1
        ClosestWallPurlin = pAbove
    Case 0
        'report the closest purlin
        If Abs(pAbove - Height) < Abs(pBelow - Height) Then
            ClosestWallPurlin = pAbove
        Else
            ClosestWallPurlin = pBelow
        End If
    End Select
End If
End Function
Function ClosestRoofPurlin(RafterLength As Variant, Optional Direction As Integer) As Double

If Direction = 1 Then
    'closest rounding up
    ClosestRoofPurlin = Application.WorksheetFunction.RoundUp(RafterLength / 60, 0) * 60
ElseIf Direction = -1 Then
    'closest rounding down
    ClosestRoofPurlin = Application.WorksheetFunction.RoundDown(RafterLength / 60, 0) * 60
Else
    ' closest without caring
    ClosestRoofPurlin = Application.WorksheetFunction.Round(RafterLength / 60, 0) * 60
End If


End Function

Private Function IsEven(PanelCount As Double) As Boolean

If (PanelCount Mod 2 = 0) = True Then
    IsEven = True
Else
    IsEven = False
End If

End Function


Private Sub EndwallPanelGen(EndwallPanels As Collection, eWall As String, b As clsBuilding, Optional FullHeightLinerPanels As Boolean)
Dim eP1 As clsPanel
Dim eP2 As clsPanel
Dim eP3 As clsPanel
Dim WainscotPanel As clsPanel
Dim WainscotFtLength As Double
Dim ePanel As clsPanel
Dim ePanelCount As Double
Dim pLengthMax As Double     'in
Dim pNum As Integer
Dim pLength As Double
Dim SpecialBottomPurlin As Boolean  ''' this boolean applies when the endwall is marked as partial or as gable only
Dim UnsplicedPanels As New Collection

Dim MaxSegments As Integer
Dim Segment1Length As Double
Dim Segment2Length As Double

Dim TopPanelLengths() As Integer
Dim FOCollection As Collection
Dim FO As clsFO

'Note: Account for Max Height term in endwall panel segment?

With b
    'determine number of panels
    ePanelCount = Application.WorksheetFunction.RoundUp(.bWidth / 3, 0)
    'top down purlins
    If .WallStatus(eWall) = "Partial" Or .WallStatus(eWall) = "Gable Only" Then SpecialBottomPurlin = True
    'Check for Wainscot
    If .Wainscot(eWall) <> "None" Then
        Set WainscotPanel = New clsPanel
        WainscotPanel.PanelLength = CDbl(Left(.Wainscot(eWall), 2))
        'only use wainscot ft length when not doing liner panels
        If FullHeightLinerPanels = False Then WainscotFtLength = WainscotPanel.PanelLength / 12
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Gable Roofs ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If .rShape = "Gable" Then
        Select Case .WallStatus(eWall)
        Case "Exclude"
            Exit Sub
        Case "Include", "Partial"
            'MaxHeight = ((.bHeight - .LengthAboveFinishedFloor(eWall)) * 12) + ((.bWidth / 2) * .rPitch)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' If Even # of panels
            If IsEven(ePanelCount) = True Then
                'add lengths symetrically to all panels
                For pNum = 1 To (ePanelCount / 2)
                    'new panel class
                    Set eP1 = New clsPanel
                    'Check if adding length to first panel or not
                    If pNum = 1 Then
                        ''''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                        If .rPitch = 1 Then
                            eP1.PanelLength = (.bHeight - .LengthAboveFinishedFloor(eWall) - WainscotFtLength) * 12 ' only bHeight for rPitch 1
                        Else
                            eP1.PanelLength = ((.bHeight - .LengthAboveFinishedFloor(eWall) - WainscotFtLength) * 12) + (.rPitch * 3) ''' bHeight + rPitch contribution
                        End If
                    ElseIf pNum <> 1 Then
                        eP1.PanelLength = pLengthMax + (.rPitch * 3)
                    End If
                    'one panel for each side of the endwall
                    eP1.Quantity = 1
                    eP1.rEdgePosition = (((.bWidth - (ePanelCount * 3)) / 2) + (pNum - 1) * 3) * 12  'add panel edge position to total wall panel width overage/2 to evenly center panel overage between corners
                    'add panel to collection
                    UnsplicedPanels.Add eP1
                    'create, add the duplicate panel for the other side of the endwall to the collection
                    Set eP1 = New clsPanel
                    eP1.Quantity = 1
                    eP1.PanelLength = UnsplicedPanels(UnsplicedPanels.Count).PanelLength
                    eP1.rEdgePosition = (((.bWidth + ((ePanelCount * 3) - .bWidth) / 2)) - (pNum * 3)) * 12     'center overage again, this time for the far side
                    UnsplicedPanels.Add eP1
                    'update running panel length
                    pLengthMax = eP1.PanelLength
                Next pNum
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' If odd # of panels
            ElseIf IsEven(ePanelCount) = False Then
                'add lengths symetrically to all panels except the long middle one
                For pNum = 1 To ((ePanelCount - 1) / 2) ''' Add symetrically to all panels but the last
                    'new panel class
                    Set eP1 = New clsPanel
                    'Check if adding length to first panel or not
                    If pNum = 1 Then
                        ''''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                        If .rPitch = 1 Then
                            eP1.PanelLength = (.bHeight - .LengthAboveFinishedFloor(eWall) - WainscotFtLength) * 12   ' only bHeight for rPitch 1
                        Else
                            eP1.PanelLength = ((.bHeight - .LengthAboveFinishedFloor(eWall) - WainscotFtLength) * 12) + (.rPitch * 3) ''' bHeight + rPitch contribution
                        End If
                    ElseIf pNum <> 1 Then
                        eP1.PanelLength = pLengthMax + (.rPitch * 3)
                    End If
                    'one panel for each side of the endwall
                    eP1.Quantity = 1
                    eP1.rEdgePosition = (((.bWidth - (ePanelCount * 3)) / 2) + (pNum - 1) * 3) * 12  'add panel edge position to total wall panel width overage/2 to evenly center panel overage between corners
                    'add panel to collection
                    UnsplicedPanels.Add eP1
                    'create, add the duplicate panel for the other side of the endwall to the collection
                    Set eP1 = New clsPanel
                    eP1.Quantity = 1
                    eP1.PanelLength = UnsplicedPanels(UnsplicedPanels.Count).PanelLength
                    eP1.rEdgePosition = (((.bWidth + ((ePanelCount * 3) - .bWidth) / 2)) - (pNum * 3)) * 12     'center overage again, this time for the far side
                    UnsplicedPanels.Add eP1
                    'update running panel length
                    pLengthMax = eP1.PanelLength
                Next pNum
                ' add roof pitch contribution again one more time for middle panel
                Set eP1 = New clsPanel
                eP1.PanelLength = pLengthMax + (.rPitch * 3)
                ' 1 because center panel
                eP1.Quantity = 1
                eP1.rEdgePosition = (((.bWidth - (ePanelCount * 3)) / 2) + (pNum * 3)) * 12    'center panel position
                'add panels to collection
                UnsplicedPanels.Add eP1
                'update running panel length
                pLengthMax = eP1.PanelLength
            End If
        Case "Gable Only"
             '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' If Even # of panels
            If IsEven(ePanelCount) = True Then
                'add lengths symetrically to all panels
                For pNum = 1 To (ePanelCount / 2)
                    ''''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                    If pNum = 1 And .rPitch <> 1 Then
                        Set eP1 = New clsPanel
                        eP1.PanelLength = .rPitch * 3   'rPitch contribution
                    ElseIf pNum <> 1 Then
                        Set eP1 = New clsPanel
                        eP1.PanelLength = pLengthMax + (.rPitch * 3)
                    End If
                    'add panels to collection
                    If Not eP1 Is Nothing Then
                        eP1.Quantity = 2
                        UnsplicedPanels.Add eP1
                        pLengthMax = eP1.PanelLength
                    End If
                Next pNum
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' If odd # of panels
            ElseIf IsEven(ePanelCount) = False Then
                'add lengths symetrically to all panels except the long middle one
                For pNum = 1 To ((ePanelCount - 1) / 2) ''' Add symetrically to all panels but the last
                    'Check if adding length to first panel or not
                    If pNum = 1 And .rPitch <> 1 Then
                        Set eP1 = New clsPanel
                        eP1.PanelLength = (.bHeight * 12) + (.rPitch * 3) 'rPitch contribution
                    ElseIf pNum <> 1 Then
                        Set eP1 = New clsPanel
                        eP1.PanelLength = pLengthMax + (.rPitch * 3)
                    End If
                    'add panels to collection, update plength
                    If Not eP1 Is Nothing Then
                        eP1.Quantity = 2
                        UnsplicedPanels.Add eP1
                        pLengthMax = eP1.PanelLength
                    End If
                Next pNum
                ' add roof pitch contribution again one more time for middle panel
                Set eP1 = New clsPanel
                eP1.PanelLength = pLengthMax + (.rPitch * 3)
                '1 for each endwall
                eP1.Quantity = 1
                'add panels to collection
                UnsplicedPanels.Add eP1
                'update running panel length
                pLengthMax = eP1.PanelLength
            End If
        End Select
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Single Slope Roofs ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf .rShape = "Single Slope" Then
        Select Case .WallStatus(eWall)
        Case "Exclude"
            Exit Sub
        Case "Include", "Partial"
            For pNum = 1 To ePanelCount
                Set eP1 = New clsPanel
                'Check if adding length to first panel or not
                If pNum = 1 Then
                    ''''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                    If .rPitch = 1 Then
                        eP1.PanelLength = (.bHeight - .LengthAboveFinishedFloor(eWall) - WainscotFtLength) * 12   ' only bHeight for rPitch 1
                    Else
                        eP1.PanelLength = ((.bHeight - .LengthAboveFinishedFloor(eWall) - WainscotFtLength) * 12) + (.rPitch * 3) ''' bHeight + rPitch contribution
                    End If
                ElseIf pNum <> 1 Then
                    eP1.PanelLength = pLengthMax + (.rPitch * 3)
                End If
                eP1.Quantity = 1
                'calculate edge position correctly since the algorithm generates panel lengths in order of ascending height. therefore, position must be reversed on e1 due to global positive direction being to the left in profile view
                If eWall = "e1" Then
                    eP1.rEdgePosition = ((ePanelCount * 3) - pNum * 3) * 12    'subtract current panel width from full width of endwall panels (including any overage). This means any overage is always taken out of the width of the shortest endwall panel
                ElseIf eWall = "e3" Then
                    eP1.rEdgePosition = ((.bWidth - (ePanelCount * 3)) + (pNum - 1) * 3) * 12
                End If
                'add panels to collection
                UnsplicedPanels.Add eP1
                pLengthMax = eP1.PanelLength
            Next pNum
        Case "Gable Only"
            For pNum = 1 To ePanelCount
                ''''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                If pNum = 1 And .rPitch <> 1 Then
                    Set eP1 = New clsPanel
                    eP1.PanelLength = .rPitch * 3 ''' bHeight + rPitch contribution
                ElseIf pNum <> 1 Then
                    Set eP1 = New clsPanel
                    eP1.PanelLength = pLengthMax + (.rPitch * 3)
                End If
                'add panel to collection
                If Not eP1 Is Nothing Then
                    eP1.Quantity = 1
                    UnsplicedPanels.Add eP1
                    pLengthMax = eP1.PanelLength
                End If
            Next pNum
        End Select
    End If
End With


'''''''''''''''''''''''''''''''''''''''''''''''''' Segment Panels & Account for FO Cutouts''''''''''''''''''''''''''''''''''''''''
'set FO Collection
If eWall = "e1" Then Set FOCollection = b.e1FOs Else Set FOCollection = b.e3FOs
''''''''''''''''''''''''''''''''''''''''''''''''''' Using the max length of the unsegmented panels, first calculate the evenly porportioned segment lengths ''''''''''''
''' Max Segments (Not Factoring FO Cutouts) = 1
If pLengthMax <= (42 * 12) Then
    MaxSegments = 1
    ' Only 1 segment and panels don't need splicing. Add panels, and exit sub
    For Each ePanel In UnsplicedPanels
        For Each FO In FOCollection
            If (FO.FOType = "OHDoor" Or FO.FOType = "MiscFO") And FO.Width >= 7 * 12 And FO.bEdgeHeight = 0 Then
                If ePanel.rEdgePosition >= (FO.rEdgePosition + 3 * 12) And ePanel.lEdgePosition <= (FO.lEdgePosition - 3 * 12) Then
                    ePanel.PanelLength = ePanel.PanelLength - (FO.Height - (WainscotFtLength * 12))
                End If
            End If
        Next FO
        'deduct 8" from full height liners
        If FullHeightLinerPanels = True Then ePanel.PanelLength = ePanel.PanelLength - 8
        'make sure panel length is > 0 before adding. this can happen when full height liners are less than 8" from the ceiling
        If ePanel.PanelLength > 0 Then EndwallPanels.Add ePanel
    Next ePanel
    'save overlaps to building class
    If eWall = "e1" Then
        b.e1WallPanelOverlaps = 0
    ElseIf eWall = "e3" Then
        b.e3WallPanelOverlaps = 0
    End If
    GoTo FinishCollection '' <--------- Exit and finish collection at the end of 1 segment condition
''' Max Segments (Not Factoring FO Cutouts) = 2
ElseIf pLengthMax <= ((79 * 12) + 3.5) Then
    MaxSegments = 2
    Segment1Length = ClosestWallPurlin(pLengthMax / 2, 0, SpecialBottomPurlin)
    'correct if greater than 37'3.5" for purlins that go from the bottom up
    If SpecialBottomPurlin = False Then If Segment1Length > ((37 * 12) + 3.5) Then Segment1Length = ((37 * 12) + 3.5)
    'add overlaps to building class
    If eWall = "e1" Then
        b.e1WallPanelOverlaps = 1
    ElseIf eWall = "e3" Then
        b.e3WallPanelOverlaps = 1
    End If
''' Max Segments (Not Factoring FO Cutouts) = 3
Else
    MaxSegments = 3
    Segment1Length = ClosestWallPurlin(pLengthMax / 3, 0, SpecialBottomPurlin)
    Segment2Length = ClosestWallPurlin(pLengthMax / 3 * 2, 0, SpecialBottomPurlin) - Segment1Length
    'add overlaps to building class
    If eWall = "e1" Then
        b.e1WallPanelOverlaps = 2
    ElseIf eWall = "e3" Then
        b.e3WallPanelOverlaps = 2
    End If
End If


''''''''' Splice using the segment lengths calculated above ''''''''''''''
''' Note: This occurs when we have an unspliced panel collection which we know must include panels which need to be segmented
For Each ePanel In UnsplicedPanels
    If ePanel.PanelLength <= Segment1Length Then
        ''' check for intersecting FOs
        For Each FO In FOCollection
            If (FO.FOType = "OHDoor" Or FO.FOType = "MiscFO") And FO.Width >= 7 * 12 And FO.bEdgeHeight = 0 Then
                If ePanel.rEdgePosition >= (FO.rEdgePosition + 3 * 12) And ePanel.lEdgePosition <= (FO.lEdgePosition - 3 * 12) Then
                    ePanel.PanelLength = ePanel.PanelLength - (FO.Height - (WainscotFtLength * 12))
                End If
            End If
        Next FO
        Set eP1 = New clsPanel
        eP1.PanelLength = ePanel.PanelLength
        'deduct 8" from full height liners
        If FullHeightLinerPanels = True Then eP1.PanelLength = eP1.PanelLength - 8
        'EndwallPanels.Add ePanel
    ElseIf MaxSegments = 2 Then
        If ePanel.PanelLength > Segment1Length Then
            ''' check for intersecting FOs
            For Each FO In FOCollection
                If (FO.FOType = "OHDoor" Or FO.FOType = "MiscFO") And FO.Width >= 7 * 12 And FO.bEdgeHeight = 0 Then
                    If ePanel.rEdgePosition >= (FO.rEdgePosition + 3 * 12) And ePanel.lEdgePosition <= (FO.lEdgePosition - 3 * 12) Then
                        'If FO takes up less than segment 1, create first panel from height remaining after cutout and create segment 2 normally (since it's above the FO and not effected)
                        If Segment1Length - (FO.Height - (WainscotFtLength * 12)) > 0 Then
                            Set eP1 = New clsPanel
                            eP1.PanelLength = Segment1Length - (FO.Height - (WainscotFtLength * 12)) + 1.5
                            Set eP2 = New clsPanel
                            eP2.PanelLength = ePanel.PanelLength - Segment1Length + 1.5
                            If FullHeightLinerPanels = True Then eP2.PanelLength = eP2.PanelLength - 8
                            GoTo AddSegmentedPanels
                        ''' If FO takes up the entirity of segment 1 or more, add segment 2 without overlap and subtract the height remaining after the FO cutout
                        ElseIf Segment1Length - (FO.Height - (WainscotFtLength * 12)) <= 0 Then
                            Set eP2 = New clsPanel
                            eP2.PanelLength = ePanel.PanelLength - (FO.Height - (WainscotFtLength * 12))
                            If FullHeightLinerPanels = True Then eP2.PanelLength = eP2.PanelLength - 8
                            GoTo AddSegmentedPanels
                        End If
                    End If
                End If
            Next FO
            '''''''''''' no intersecting FOs
            Set eP1 = New clsPanel
            eP1.PanelLength = Segment1Length + 1.5
            Set eP2 = New clsPanel
            eP2.PanelLength = ePanel.PanelLength - Segment1Length + 1.5
            If FullHeightLinerPanels = True Then eP2.PanelLength = eP2.PanelLength - 8
        End If
    ElseIf MaxSegments = 3 Then
        If ePanel.PanelLength <= (Segment1Length + Segment2Length) Then
            ''' check for intersecting FOs
            For Each FO In FOCollection
                If (FO.FOType = "OHDoor" Or FO.FOType = "MiscFO") And FO.Width >= 7 * 12 And FO.bEdgeHeight = 0 Then
                    If ePanel.rEdgePosition >= (FO.rEdgePosition + 3 * 12) And ePanel.lEdgePosition <= (FO.lEdgePosition - 3 * 12) Then
                        'If FO takes up less than segment 1, create first panel from height remaining after cutout and create segment 2 normally (since it's above the FO and not effected)
                        If Segment1Length - (FO.Height - (WainscotFtLength * 12)) > 0 Then
                            Set eP1 = New clsPanel
                            eP1.PanelLength = Segment1Length - (FO.Height - (WainscotFtLength * 12)) + 1.5
                            Set eP2 = New clsPanel
                            eP2.PanelLength = ePanel.PanelLength - Segment1Length + 1.5
                            GoTo AddSegmentedPanels
                        ''' If FO takes up the entirity of segment 1 or more, add segment 2 without overlap and subtract the height remaining after the FO cutout
                        ElseIf Segment1Length - (FO.Height - (WainscotFtLength * 12)) <= 0 Then
                            Set eP2 = New clsPanel
                            eP2.PanelLength = (Segment2Length + Segment1Length) - (FO.Height - (WainscotFtLength * 12)) + 1.5
                            GoTo AddSegmentedPanels
                        End If
                    End If
                End If
            Next FO
            '''''''''''' no intersecting FOs
            Set eP1 = New clsPanel
            eP1.PanelLength = Segment1Length + 1.5
            Set eP2 = New clsPanel
            eP2.PanelLength = ePanel.PanelLength - Segment1Length + 1.5
        ElseIf ePanel.PanelLength > (Segment1Length + Segment2Length) Then
            ''' check for intersecting FOs
            For Each FO In FOCollection
                If (FO.FOType = "OHDoor" Or FO.FOType = "MiscFO") And FO.Width >= 7 * 12 And FO.bEdgeHeight = 0 Then
                    If ePanel.rEdgePosition >= (FO.rEdgePosition + 3 * 12) And ePanel.lEdgePosition <= (FO.lEdgePosition - 3 * 12) Then
                        'If FO takes up less than segment 1, create first panel from height remaining after cutout and create segment 2 normally (since it's above the FO and not effected)
                        If Segment1Length - (FO.Height - (WainscotFtLength * 12)) > 0 Then
                            Set eP1 = New clsPanel
                            eP1.PanelLength = Segment1Length - (FO.Height - (WainscotFtLength * 12)) + 1.5
                            Set eP2 = New clsPanel
                            eP2.PanelLength = Segment2Length + 3
                            Set eP3 = New clsPanel
                            eP3.PanelLength = ePanel.PanelLength - Segment1Length - Segment2Length + 1.5
                            If FullHeightLinerPanels = True Then eP3.PanelLength = eP3.PanelLength - 8
                            GoTo AddSegmentedPanels
                        ''' If FO takes up the entirity of segment 1 or more, add segment 2 without overlap and subtract the height remaining after the FO cutout
                        ElseIf Segment1Length - (FO.Height - (WainscotFtLength * 12)) <= 0 Then
                            Set eP2 = New clsPanel
                            eP2.PanelLength = (Segment2Length + Segment1Length) - (FO.Height - (WainscotFtLength * 12)) + 1.5
                            Set eP3 = New clsPanel
                            eP3.PanelLength = ePanel.PanelLength - Segment1Length - Segment2Length + 1.5
                            If FullHeightLinerPanels = True Then eP3.PanelLength = eP3.PanelLength - 8
                            GoTo AddSegmentedPanels
                        End If
                    End If
                End If
            Next FO
            '''''''''''' no intersecting FOs
            Set eP1 = New clsPanel
            eP1.PanelLength = Segment1Length + 1.5
            Set eP2 = New clsPanel
            eP2.PanelLength = Segment2Length + 3
            Set eP3 = New clsPanel
            eP3.PanelLength = ePanel.PanelLength - Segment1Length - Segment2Length + 1.5
            'deduct 8" from full height liners
            If FullHeightLinerPanels = True Then eP3.PanelLength = eP3.PanelLength - 8
        End If
    End If
AddSegmentedPanels:
    '''''''''''''''add segmented panels to collection
    If Not eP1 Is Nothing Then
        eP1.Quantity = ePanel.Quantity
        If eP1.PanelLength > 0 Then EndwallPanels.Add eP1
        Set eP1 = Nothing
    End If
    If Not eP2 Is Nothing Then
        eP2.Quantity = ePanel.Quantity
        If eP2.PanelLength > 0 Then EndwallPanels.Add eP2
        Set eP2 = Nothing
    End If
    If Not eP3 Is Nothing Then
        eP3.Quantity = ePanel.Quantity
        If eP3.PanelLength > 0 Then EndwallPanels.Add eP3
        Set eP3 = Nothing
    End If
Next ePanel

'''


FinishCollection:

'add parameters
For Each ePanel In EndwallPanels
    ePanel.PanelShape = b.wPanelShape
    ePanel.PanelType = b.wPanelType
    ePanel.PanelColor = b.wPanelColor
    ePanel.PanelMeasurement = ImperialMeasurementFormat(ePanel.PanelLength)
Next ePanel

If Not WainscotPanel Is Nothing And FullHeightLinerPanels = False Then
    WainscotPanel.PanelMeasurement = ImperialMeasurementFormat(WainscotPanel.PanelLength)
    WainscotPanel.Quantity = Application.WorksheetFunction.RoundUp(b.bWidth / 3, 0)
    WainscotPanel.PanelColor = EstSht.Range(eWall & "_Wainscot").offset(0, 2).Value
    WainscotPanel.PanelType = EstSht.Range(eWall & "_Wainscot").offset(0, 1).Value
    WainscotPanel.PanelShape = b.wPanelShape
    EndwallPanels.Add WainscotPanel
End If

Call DuplicateMaterialRemoval(EndwallPanels, "Panel")


End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Trim calculation (and downspouts and gutters)
Private Sub TrimPieceCalc(ByRef TrimCollection As Collection, NetTrimLength As Double, TrimType As String, Optional rPitchString As String, Optional DownspoutQty As Integer, Optional b As clsBuilding)
Dim Trim As clsTrim
Dim ExistingTrim As clsTrim

'vars to fill in Trim class
Dim tQty As Integer
Dim tLength As Double
Dim tTypeString As String
Dim RemainingLength As Double
Dim t As Integer
Dim DuplicateFound As Boolean
Dim LargestTrimDivisor As Integer
Dim IdealTrimSize As Integer
Dim TrimSegmentsRequired As Integer
Dim n As Integer
Dim tLengthRemaining As Integer
Dim CurrentLength As Integer


'trim Type
Select Case TrimType
Case "Rake"
    tTypeString = "Rake Trim"
    'check for trim lengths
    With b
        '''single slope
        If .rShape = "Single Slope" Then
            'sidewall 2
            If .s2RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength <= 244 Then
                'add trim for single side
                Set Trim = New clsTrim
                Trim.tLength = NearestTrimSize(.s2RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength, 1, , True)
                Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength)
                Trim.Quantity = 2
                Trim.tType = tTypeString
                TrimCollection.Add Trim
                'decrease needed trim length
                NetTrimLength = NetTrimLength - (Trim.tLength * Trim.Quantity)
            End If
        ''''Gable
        ElseIf .rShape = "Gable" Then
            'sidewall 2
            If .s2RafterSheetLength + .s2ExtensionRafterLength <= 244 Then
                'add trim for single side
                Set Trim = New clsTrim
                Trim.tLength = NearestTrimSize(.s2RafterSheetLength + .s2ExtensionRafterLength, 1, , True)
                Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength)
                Trim.Quantity = 2
                Trim.tType = tTypeString
                TrimCollection.Add Trim
                'decrease needed trim length
                NetTrimLength = NetTrimLength - (Trim.tLength * Trim.Quantity)
            End If
            If .s4RafterSheetLength + .s4ExtensionRafterLength <= 244 Then
                'add trim for single side
                Set Trim = New clsTrim
                Trim.tLength = NearestTrimSize(.s4RafterSheetLength + .s4ExtensionRafterLength, 1, , True)
                Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength)
                Trim.Quantity = 2
                Trim.tType = tTypeString
                TrimCollection.Add Trim
                'decrease needed trim length
                NetTrimLength = NetTrimLength - (Trim.tLength * Trim.Quantity)
            End If
        End If
    End With
Case "Short Eave"
    tTypeString = "Short Eave" & " " & rPitchString
Case "High Eave"
    tTypeString = "High-Side Eave" & " " & rPitchString
Case "Outside Corner"
    tTypeString = "Outside Corner Trim"
    'check for single piece
    With b
        If (.bHeight * 12) <= 244 Then
            'add trim for single side
            Set Trim = New clsTrim
            Trim.tLength = NearestTrimSize(.bHeight * 12, 1, , True)
            Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength)
            If .rShape = "Gable" Then
                Trim.Quantity = 4
            ElseIf .rShape = "Single Slope" Then
                Trim.Quantity = 2
            End If
            Trim.tType = tTypeString
            TrimCollection.Add Trim
            'decrease needed trim length
            NetTrimLength = NetTrimLength - (Trim.tLength * Trim.Quantity)
        End If
        'check high side height if single slope
        If .rShape = "Single Slope" Then
            If .HighSideEaveHeight <= 244 Then
                 'add trim for single side
                Set Trim = New clsTrim
                Trim.tLength = NearestTrimSize(.HighSideEaveHeight, 1, , True)
                Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength)
                Trim.Quantity = 2
                Trim.tType = tTypeString
                TrimCollection.Add Trim
                'decrease needed trim length
                NetTrimLength = NetTrimLength - (Trim.tLength * Trim.Quantity)
            End If
        End If
    End With
Case "Base"
    tTypeString = "Base Trim"
Case "Gutter"
    tTypeString = "Sculptured Gutter Hang-On" & " " & rPitchString
Case "Downspout"
    tTypeString = "Square Downspout W/O Kickout"
Case "Jamb"
    tTypeString = "Jamb Trim"
    'Any Jamb Trim calculated here is for soffits, so calculate if only a single piece can be used
    'check for trim lengths
    With b
        'sidewall 2
        If .RafterLength <= 244 Then
            'add trim for single side
            Set Trim = New clsTrim
            Trim.tLength = NearestTrimSize(.RafterLength, 1, "Jamb", True)
            Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength)
            Trim.Quantity = 2
            Trim.tType = tTypeString
            TrimCollection.Add Trim
            'decrease needed trim length
            NetTrimLength = NetTrimLength - (Trim.tLength * Trim.Quantity)
        End If
        'sidewall 4
        If .rShape = "Gable" Then
            If .RafterLength <= 244 Then
                'add trim for single side
                Set Trim = New clsTrim
                Trim.tLength = NearestTrimSize(.RafterLength, 1, "Jamb", True)
                Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength)
                Trim.Quantity = 2
                Trim.tType = tTypeString
                TrimCollection.Add Trim
                'decrease needed trim length
                NetTrimLength = NetTrimLength - (Trim.tLength * Trim.Quantity)
            End If
        End If
    End With
Case "Head"
    tTypeString = "Head Trim W/O Kickout"
Case "Outside Angle"
    tTypeString = "2x6 Outside Angle Trim"
Case "Inside Angle"
    tTypeString = "2x8 Inside Angle Trim"
Case "Standard Wainscot"
    tTypeString = "Standard Wainscot Trim"
Case "Masonry Wainscot"
    tTypeString = "Masonry Wainscot Trim"
End Select
'exit sub if no remaining trim length
If NetTrimLength <= 0 Then
    Call DuplicateMaterialRemoval(TrimCollection, "Trim")
    Exit Sub
End If

''' note: need to add a check to see if we should start with 20'4" trim or not (in cases where trim is shorter)
Set Trim = New clsTrim
''''''' Check for starting piece size
Select Case NetTrimLength
Case Is >= 244
    LargestTrimDivisor = 240
    Trim.tMeasurement = "20'4"""
    Trim.tLength = 242
Case Is >= 218
    LargestTrimDivisor = 216
    Trim.tMeasurement = "18'2"""
    Trim.tLength = 216
Case Is >= 194
    LargestTrimDivisor = 192
    Trim.tMeasurement = "16'2"""
    Trim.tLength = 192
Case Is >= 170
    LargestTrimDivisor = 168 '
    Trim.tMeasurement = "14'2"""
    Trim.tLength = 168
Case Is >= 146
    LargestTrimDivisor = 144
    Trim.tMeasurement = "12'2"""
    Trim.tLength = 144
Case Else
    LargestTrimDivisor = 120
    Trim.tMeasurement = "10'2"""
    Trim.tLength = 120
End Select

'check for pieces of the largest size trim
tQty = Application.WorksheetFunction.RoundDown((NetTrimLength / LargestTrimDivisor), 0)
If TrimType <> "Downspout" Then
    Trim.Quantity = tQty
ElseIf TrimType = "Downspout" Then
    Trim.Quantity = tQty * DownspoutQty
End If
Trim.tType = tTypeString
Trim.clsType = "Trim"
' add trim to collection
TrimCollection.Add Trim
'''find other trim size
Set Trim = New clsTrim
Trim.tType = tTypeString
'find remaining length
RemainingLength = NetTrimLength - (LargestTrimDivisor * tQty)
If RemainingLength <> 0 Then
    'find size, write length and qty to class
    If TrimType <> "Jamb" And TrimType <> "Head" Then
        Trim.tMeasurement = NearestTrimSize(RemainingLength, 1)
        Trim.tLength = NearestTrimSize(RemainingLength, 1, , True)
    ElseIf TrimType = "Jamb" Then
        Trim.tMeasurement = NearestTrimSize(RemainingLength, 1, "Jamb")
        Trim.tLength = NearestTrimSize(RemainingLength, 1, "Jamb", True)
    ElseIf TrimType = "Head" Then
        Trim.tMeasurement = NearestTrimSize(RemainingLength, 1, "Head")
        Trim.tLength = NearestTrimSize(RemainingLength, 1, "Head", True)
    End If
    'just increase quantity of the corresponding 20'4" (or largest starting size) if the remaining trim size rounds to it
    For t = 1 To TrimCollection.Count
        If (Trim.tMeasurement = TrimCollection(t).tMeasurement) And (Trim.tType = TrimCollection(t).tType) Then
            If TrimType <> "Downspout" Then
                TrimCollection(t).Quantity = TrimCollection(t).Quantity + 1
            ElseIf TrimType = "Downspout" Then
                TrimCollection(t).Quantity = TrimCollection(t).Quantity + DownspoutQty
            End If
            'mark duplicate as found, exit
            DuplicateFound = True
            Exit For
        End If
    Next t
    'if no duplicate found add to collection
    If DuplicateFound = False Then
        If TrimType <> "Downspout" Then
            Trim.Quantity = 1
        ElseIf TrimType = "Downspout" Then
            Trim.Quantity = DownspoutQty
        End If
        Trim.Quantity = 1
        Trim.clsType = "Trim"
        TrimCollection.Add Trim
    End If
    
End If
    


End Sub


'' function returns string of the nearest available rake trim size
Function NearestTrimSize(Length As Variant, Optional Direction As Integer, Optional UniqueTrimType As String, Optional NumericOutput As Boolean) As Variant

'DESCRIPTION: Function returns the nearest value to a target
'INPUT: Pass the function a range of cells, a target value that you want to find a number closest to
' and an optional direction variable described below.
'OPTIONS: Set the optional variable Direction equal to 0 or blank to find the closest value
' Set equal to -1 to find the closest value below your target
' set equal to 1 to find the closest value above your target
'OUTPUT: The output is the number in the range closest to your target value.
' Because the output is a variant, the address of the closest number can also be returned when
' calling this function from another VBA macro.
Dim t As Variant
Dim u As Variant
Dim Trims() As Variant
Dim Trim As Variant
Dim tSize As Variant
Dim NearestTrimSizeString As String

'check for trim type
If UniqueTrimType = "" Then
    Trims = Array(122, 146, 170, 194, 218, 244)
ElseIf UniqueTrimType = "Head" Then
    Trims = Array(42, 75, 87, 99, 122, 123, 147, 171, 195, 219, 244)
ElseIf UniqueTrimType = "Jamb" Then
    Trims = Array(86, 122, 146, 170, 194, 218, 244)
End If

t = 1.79769313486231E+308 'initialize
For Each Trim In Trims
    If IsNumeric(Trim) Then
        u = Abs(Trim - Length)
        If Direction > 0 And Trim >= Length Then
            'only report if closer number is greater than the target
            If u < t Then
                t = u
                tSize = Trim
            End If
        ElseIf Direction < 0 And Trim <= Length Then
            'only report if closer number is less than the target
            If u < t Then
                t = u
                tSize = Trim
            End If
        ElseIf Direction = 0 Then
            If u < t Then
                t = u
                tSize = Trim
            End If
        End If
    End If
Next


'return available trim name
Select Case tSize
Case 42
    NearestTrimSizeString = "3'6"""
Case 75
    NearestTrimSizeString = "6'3"""
Case 86
    NearestTrimSizeString = "7'2"""
Case 87
    NearestTrimSizeString = "7'3"""
Case 99
    NearestTrimSizeString = "8'3"""
Case 122
    NearestTrimSizeString = "10'2"""
Case 123
    NearestTrimSizeString = "10'3"""
Case 146
    NearestTrimSizeString = "12'2"""
Case 147
    NearestTrimSizeString = "12'3"""
Case 170
    NearestTrimSizeString = "14'2"""
Case 171
    NearestTrimSizeString = "14'3"""
Case 194
    NearestTrimSizeString = "16'2"""
Case 195
    NearestTrimSizeString = "16'3"""
Case 218
    NearestTrimSizeString = "18'2"""
Case 219
    NearestTrimSizeString = "18'3"""
Case 244
    NearestTrimSizeString = "20'4"""
End Select

'output
If NumericOutput = False Then
    NearestTrimSize = NearestTrimSizeString
ElseIf NumericOutput = True Then
    NearestTrimSize = tSize
End If

    

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sub for generating framed opening materials
Private Sub FOMaterialGen(MatSht As Worksheet, TrimCollection As Collection, MiscCollection As Collection)
Dim FOCell As Range
Dim hTrim As clsTrim
Dim jTrim As clsTrim
Dim CanopyQty As Integer
Dim DoorSlabNoGlassQty As Integer
Dim DoorSlabGlassQty As Integer
Dim PDoorMaterials As Collection
Dim OHDoorMaterials As Collection
Dim WindowMaterials As Collection
Dim MiscFOMaterials As Collection
Dim PDoors As Collection
Dim OHDoors As Collection
Dim Windows As Collection
Dim MiscFOs As Collection
Dim FO As clsFO
Dim m As Integer
Dim m2 As Integer
Dim WriteCell As Range
Dim FOMaterial As clsTrim
Dim FOTrimColor As String
Dim FOWidth As Integer
Dim FOHeight As Integer
Dim RemainingWidth As Integer
Dim RemainingHeight As Integer
Dim HeadTrimMeasurement As String
Dim HeadTrimLength As Integer
'Canopy

Dim q As Integer
Dim CombinedLength As Integer
Dim tType As String
Dim NewQuantity As Integer


'new material collections
Set PDoorMaterials = New Collection
Set OHDoorMaterials = New Collection
Set WindowMaterials = New Collection
Set MiscFOMaterials = New Collection
'new framed opening collections
Set PDoors = New Collection
Set OHDoors = New Collection
Set Windows = New Collection
Set MiscFOs = New Collection

'Personnel Door vars
Dim DoorSize As String
'' Even though these aren't trim, trim class is used for convenience
Dim JambKit As clsTrim
Dim DoorSlab As clsTrim


With EstSht
    'set trim color
    FOTrimColor = .Range("FO_tColor").Value
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Personnel Doors
    For Each FOCell In Range(.Range("pDoorCell1"), .Range("pDoorCell12"))
        'if cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'new FO class
            Set FO = New clsFO
            FO.FOType = "PDoor"
            FO.Height = 7 * 12
            DoorSize = FOCell.offset(0, 1).Value
            'based off of size, select trim measurement appropriately
            If DoorSize = "3070" Then
                FO.Width = (3 * 12) + 3
            ElseIf DoorSize = "4070" Then
                FO.Width = (4 * 12) + 3
            End If
            '''door slab
            'with glass
            If FOCell.offset(0, 3).Value = "Yes" Then
                Set DoorSlab = New clsTrim
                'deadbolt
                If FOCell.offset(0, 6).Value = "Yes" Then
                    DoorSlab.tType = "Door Slab W/ Deadbolt, W/ Glass" & " - " & DoorSize
                Else
                    DoorSlab.tType = "Door Slab W/O Deadbolt, W/ Glass" & " - " & DoorSize
                End If
            'without glass
            Else
                Set DoorSlab = New clsTrim
                'deadbolt
                If FOCell.offset(0, 6).Value = "Yes" Then
                    DoorSlab.tType = "Door Slab W/ Deadbolt, W/O Glass" & " - " & DoorSize
                Else
                    DoorSlab.tType = "Door Slab W/O Deadbolt, W/O Glass" & " - " & DoorSize
                End If
            End If
            DoorSlab.tMeasurement = "N/A"
            DoorSlab.Quantity = 1
            DoorSlab.Color = "N/A"
            PDoorMaterials.Add DoorSlab
            '''canopy
            'jamb kit
            Set JambKit = New clsTrim
            JambKit.Quantity = 1
            JambKit.tMeasurement = FOCell.offset(0, 5).Value
            JambKit.Color = "N/A"
            'deadbolt
            If FOCell.offset(0, 6).Value = "Yes" Then
                JambKit.tType = "Jamb W/ Deadbolt" & " - " & DoorSize
            Else
                JambKit.tType = "Jamb W/O Deadbolt" & " - " & DoorSize
            End If
            'add materials to collection
            PDoorMaterials.Add JambKit
            'add FO to collection
            PDoors.Add FO
        End If
    Next FOCell
    
    '' generate Personnel Door Trim
    If PDoors.Count <> 0 Then Call OptimalFOTrimGen(PDoorMaterials, PDoors, "PDoors")
    
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Overhead
    For Each FOCell In Range(.Range("OHDoorCell1"), .Range("OHDoorCell12"))
        'if cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            Set FO = New clsFO
            FO.FOType = "OHDoor"
            'door measurements
            FOWidth = FOCell.offset(0, 1).Value * 12
            FOHeight = FOCell.offset(0, 2).Value * 12
            'fo class info
            FO.Height = FOHeight
            FO.Width = FOWidth + 3
            OHDoors.Add FO
        End If
    Next FOCell
     '' generate Overhead Door Trim
    If OHDoors.Count <> 0 Then Call OptimalFOTrimGen(OHDoorMaterials, OHDoors, "OHDoors")
    
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Windows
    For Each FOCell In Range(.Range("WindowCell1"), .Range("WindowCell12"))
        'if cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            Set FO = New clsFO
            FO.FOType = "Window"
            'window measurements (in inches already)
            FOWidth = FOCell.offset(0, 1).Value + 3
            FOHeight = FOCell.offset(0, 2).Value
            FO.Width = FOWidth
            FO.Height = FOHeight
            'add window to FO collection
            Windows.Add FO
        End If
    Next FOCell
    
    '' generate Window Trim
    If Windows.Count <> 0 Then Call OptimalFOTrimGen(WindowMaterials, Windows, "Windows")
    
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Misc FO
    For Each FOCell In Range(.Range("MiscFOCell1"), .Range("MiscFOCell12"))
        'if cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            Set FO = New clsFO
            FO.FOType = "MiscFO"
            'Misc FO lengths
            FOWidth = (FOCell.offset(0, 1).Value * 12) + 3
            FOHeight = FOCell.offset(0, 2).Value * 12
            'add lengths to FO class
            FO.Width = FOWidth
            FO.Height = FOHeight
            'add FO to MiscFO collection
            MiscFOs.Add FO
        End If
    Next FOCell
    
    '' generate Misc FO Trim
    If MiscFOs.Count <> 0 Then Call OptimalFOTrimGen(MiscFOMaterials, MiscFOs, "MiscFOs")
    
End With


'remove duplicate materials, combine
Call DuplicateMaterialRemoval(PDoorMaterials, "Trim")
Call DuplicateMaterialRemoval(OHDoorMaterials, "Trim")
Call DuplicateMaterialRemoval(WindowMaterials, "Trim")
Call DuplicateMaterialRemoval(MiscFOMaterials, "Trim")

Call TrimCombine(OHDoorMaterials)
Call TrimCombine(WindowMaterials)
Call TrimCombine(MiscFOMaterials)

                           
'write
With MatSht
    'if no trim, then delete all headings and exit sub
    If PDoorMaterials.Count = 0 And OHDoorMaterials.Count = 0 And WindowMaterials.Count = 0 And MiscFOMaterials.Count = 0 Then
        Range(.Range("PDoorMatQtyCell1").offset(-4, 0), .Range("MiscFOMatQtyCell1").offset(2, 0)).EntireRow.Delete
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Personnel Doors
    If PDoorMaterials.Count = 0 Then
        .Range("PDoorMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("PDoorMatQtyCell1")
        'go through material collection
        For Each FOMaterial In PDoorMaterials
            'insert new row if not the first write cell in the section
            If WriteCell <> .Range("PDoorMatQtyCell1") Then .Rows(WriteCell.Row + 1).Insert
            Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge
            'add materials
            WriteCell.Value = FOMaterial.Quantity
            WriteCell.offset(0, 1).Value = FOMaterial.tType
            WriteCell.offset(0, 3).Value = FOMaterial.tMeasurement
            WriteCell.offset(0, 4).Value = FOMaterial.Color
            'update write cell
            Set WriteCell = WriteCell.offset(1, 0)
        Next FOMaterial
        '''Canopys
        If CanopyQty <> 0 Then
           .Rows(WriteCell.Row + 1).Insert
            Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge
            WriteCell.Value = CanopyQty
            WriteCell.offset(0, 1).Value = "Canopy"
            WriteCell.offset(0, 3).Value = "N/A"
            WriteCell.offset(0, 4).Value = "N/A"
            Set WriteCell = WriteCell.offset(1, 0)
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Overhead Doors
    If OHDoorMaterials.Count = 0 Then
        .Range("OHDoorMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("OHDoorMatQtyCell1")
        'go through material collection
        For Each FOMaterial In OHDoorMaterials
            'insert new row if not the first write cell in the section
            If WriteCell.Address <> .Range("OHDoorMatQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
            Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge
            'add materials
            WriteCell.Value = FOMaterial.Quantity
            WriteCell.offset(0, 1).Value = FOMaterial.tType
            WriteCell.offset(0, 3).Value = FOMaterial.tMeasurement
            WriteCell.offset(0, 4).Value = FOMaterial.Color
            'update write cell
            Set WriteCell = WriteCell.offset(1, 0)
        Next FOMaterial
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Windows
    If WindowMaterials.Count = 0 Then
        .Range("WindowMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("WindowMatQtyCell1")
        'go through material collection
        For Each FOMaterial In WindowMaterials
            'insert new row if not the first write cell in the section
            If WriteCell.Address <> .Range("WindowMatQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
            Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge
            'add materials
            WriteCell.Value = FOMaterial.Quantity
            WriteCell.offset(0, 1).Value = FOMaterial.tType
            WriteCell.offset(0, 3).Value = FOMaterial.tMeasurement
            WriteCell.offset(0, 4).Value = FOMaterial.Color
            'update write cell
            Set WriteCell = WriteCell.offset(1, 0)
        Next FOMaterial
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Misc FOs
    If MiscFOMaterials.Count = 0 Then
        .Range("MiscFOMatQtyCell1").offset(-2, 0).Resize(4, 1).EntireRow.Delete
    Else
        Set WriteCell = .Range("MiscFOMatQtyCell1")
        'go through material collection
        For Each FOMaterial In MiscFOMaterials
            'insert new row if not the first write cell in the section
            If WriteCell.Address <> .Range("MiscFOMatQtyCell1").Address Then .Rows(WriteCell.Row + 1).Insert
            Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge
            'add materials
            WriteCell.Value = FOMaterial.Quantity
            WriteCell.offset(0, 1).Value = FOMaterial.tType
            WriteCell.offset(0, 3).Value = FOMaterial.tMeasurement
            WriteCell.offset(0, 4).Value = FOMaterial.Color
            'update write cell
            Set WriteCell = WriteCell.offset(1, 0)
        Next FOMaterial
    End If
End With

'just add it all to trim material collection for now
For Each FOMaterial In PDoorMaterials
    If InStr(1, FOMaterial.tType, "Head Trim") <> 0 Or FOMaterial.tType = "Jamb Trim" Then
        FOMaterial.tShape = "R-Loc"
    Else
        'door slaps, windows, canpoies, etc.
        FOMaterial.tShape = "N/A"
    End If
    TrimCollection.Add FOMaterial
Next FOMaterial
For Each FOMaterial In OHDoorMaterials
    If InStr(1, FOMaterial.tType, "Head Trim") <> 0 Or FOMaterial.tType = "Jamb Trim" Then FOMaterial.tShape = "R-Loc"
    TrimCollection.Add FOMaterial
Next FOMaterial
For Each FOMaterial In WindowMaterials
    If InStr(1, FOMaterial.tType, "Head Trim") <> 0 Or FOMaterial.tType = "Jamb Trim" Then FOMaterial.tShape = "R-Loc"
    TrimCollection.Add FOMaterial
Next FOMaterial
For Each FOMaterial In MiscFOMaterials
    If InStr(1, FOMaterial.tType, "Head Trim") <> 0 Or FOMaterial.tType = "Jamb Trim" Then FOMaterial.tShape = "R-Loc"
    TrimCollection.Add FOMaterial
Next FOMaterial
End Sub


Sub DuplicateMaterialRemoval(ByRef MaterialCollection As Collection, Optional CollectionType As String)
Dim m As Integer
Dim m2 As Integer

If CollectionType = "Trim" Then
    For m = 1 To MaterialCollection.Count
        'check that not already flagged
        If MaterialCollection(m).DeleteFlag = False And MaterialCollection(m).clsType = "Trim" Then
            'check for duplicate measurements
            For m2 = 1 To MaterialCollection.Count
                If MaterialCollection(m).clsType = MaterialCollection(m2).clsType Then
                    'check that not the same material, not flagged for deletion, and duplicate measurement
                    If m2 <> m And (MaterialCollection(m).tType = MaterialCollection(m2).tType) And (MaterialCollection(m).Color = MaterialCollection(m2).Color) And _
                    MaterialCollection(m2).DeleteFlag = False And (MaterialCollection(m2).tMeasurement = MaterialCollection(m).tMeasurement) Then
                        'add quantity to existing class
                        MaterialCollection(m).Quantity = MaterialCollection(m).Quantity + MaterialCollection(m2).Quantity
                        'flag duplicate for deletion
                        MaterialCollection(m2).DeleteFlag = True
                    End If
                End If
            Next m2
        End If
    Next m
ElseIf CollectionType = "Panel" Then
    For m = 1 To MaterialCollection.Count
        'check that not already flagged
        If MaterialCollection(m).DeleteFlag = False Then
            'check for duplicate measurements
            For m2 = 1 To MaterialCollection.Count
                'check that not the same material, not flagged for deletion, and duplicate measurement
                If m2 <> m And (MaterialCollection(m).PanelType = MaterialCollection(m2).PanelType) And (MaterialCollection(m).PanelColor = MaterialCollection(m2).PanelColor) And _
                MaterialCollection(m2).DeleteFlag = False And (MaterialCollection(m2).PanelMeasurement = MaterialCollection(m).PanelMeasurement) Then
                    'add quantity to existing class
                    MaterialCollection(m).Quantity = MaterialCollection(m).Quantity + MaterialCollection(m2).Quantity
                    'flag duplicate for deletion
                    MaterialCollection(m2).DeleteFlag = True
                End If
            Next m2
        End If
    Next m
ElseIf CollectionType = "Misc" Then
    For m = 1 To MaterialCollection.Count
        'check that not already flagged
        If MaterialCollection(m).DeleteFlag = False Then
            'check for duplicate measurements
            For m2 = 1 To MaterialCollection.Count
                'check that not the same material, not flagged for deletion, and duplicate measurement
                If m2 <> m And (MaterialCollection(m).Name = MaterialCollection(m2).Name) And (MaterialCollection(m).Color = MaterialCollection(m2).Color) And _
                MaterialCollection(m2).DeleteFlag = False And (MaterialCollection(m2).Measurement = MaterialCollection(m).Measurement) Then
                    'add quantity to existing class
                    MaterialCollection(m).Quantity = MaterialCollection(m).Quantity + MaterialCollection(m2).Quantity
                    'flag duplicate for deletion
                    MaterialCollection(m2).DeleteFlag = True
                End If
            Next m2
        End If
    Next m
ElseIf CollectionType = "Steel" Then
    For m = 1 To MaterialCollection.Count
        'Check that not already flagged
        If MaterialCollection(m).DeleteFlag = False Then
            'check for duplicate measurements
            For m2 = 1 To MaterialCollection.Count
                If MaterialCollection(m).clsType = "Member" And MaterialCollection(m2).clsType = "Member" Then
                'check that not the same material, not flagged for deletion, and duplicate measurement
                    If m2 <> m And (MaterialCollection(m).Length = MaterialCollection(m2).Length) And _
                    (MaterialCollection(m).Size = MaterialCollection(m2).Size) And MaterialCollection(m2).clsType = "Member" And MaterialCollection(m2).DeleteFlag = False Then
                        'Add quatities
                        'Debug.Print vbNewLine & "m - " & m & " - " & MaterialCollection(m).Length & " - " & MaterialCollection(m).DeleteFlag & " - " & MaterialCollection(m).Placement
                        'Debug.Print "m2 - " & m2 & " - " & MaterialCollection(m2).Length & " - " & MaterialCollection(m2).DeleteFlag & " - " & MaterialCollection(m2).Placement
                        
                        MaterialCollection(m).Qty = MaterialCollection(m).Qty + MaterialCollection(m2).Qty
                        'flag duplicate for deletion
                        MaterialCollection(m2).DeleteFlag = True
                    End If
                End If
            Next m2
        End If
    Next m
End If

For m = MaterialCollection.Count To 1 Step -1
    If MaterialCollection(m).DeleteFlag = True Then
        MaterialCollection.Remove m
    End If
Next m

End Sub

Private Sub TrimCombine(ByRef MaterialCollection As Collection)
Dim m As Integer
Dim tType As String
Dim CombinedLength As Integer
Dim FOMaterial As clsTrim
Dim q As Integer


For m = 1 To MaterialCollection.Count
    With MaterialCollection(m)
        'skip materials other than trim
        If InStr(1, .tType, "Door") <> 0 Or InStr(1, .tType, "Deadbolt") <> 0 Or InStr(1, .tType, "Canopy") <> 0 Then GoTo NextMaterial
        ''''''''''''''''''FOR NOW, Only combine 10'2 pieces of trim
        If .tLength <> 122 Then GoTo NextMaterial           ''''''''''''''''''FOR NOW, Only combine 10'2 pieces of trim
        'check that trim can be combined
        If .Quantity > 1 And .tLength <= 122 Then
            'flag trim piece for deletion
            .DeleteFlag = True
            'find trim type
            If InStr(1, .tType, "Head") <> 0 Then
                tType = "Head"
            ElseIf InStr(1, .tType, "Jamb") <> 0 Then
                tType = "Jamb"
            Else
                tType = ""
            End If
            'reset combined length
            CombinedLength = 0
            For q = 1 To .Quantity
                'check if can add to previous piece and be less than 20'4"
                If (CombinedLength + .tLength) <= 244 Then
                    'combine with previous trim piece
                    CombinedLength = CombinedLength + .tLength
                Else
                    'add a new piece with the new trim length
                    Set FOMaterial = New clsTrim
                    FOMaterial.Color = .Color
                    FOMaterial.tType = .tType
                    FOMaterial.tMeasurement = NearestTrimSize(CombinedLength, 1, tType)
                    FOMaterial.tLength = NearestTrimSize(CombinedLength, 1, tType, True)
                    FOMaterial.Quantity = 1
                    MaterialCollection.Add FOMaterial
                    'reset combined to current piece
                    CombinedLength = 0 + .tLength
                End If
            Next q
            'If left over material, add to collection
            If CombinedLength <> 0 Then
                Set FOMaterial = New clsTrim
                FOMaterial.Color = .Color
                FOMaterial.tType = .tType
                FOMaterial.tMeasurement = NearestTrimSize(CombinedLength, 1, tType)
                FOMaterial.tLength = NearestTrimSize(CombinedLength, 1, tType, True)
                FOMaterial.Quantity = 1
                MaterialCollection.Add FOMaterial
            End If
        End If
    End With
NextMaterial:
Next m

'remove duplicate material
Call DuplicateMaterialRemoval(MaterialCollection, "Trim")

End Sub

Private Sub OptimalFOTrimGen(ByRef MaterialCollection As Collection, ByRef FOs As Collection, FOType As String)
Dim FO As clsFO
Dim SplitFO As clsFO
Dim CombinedWidth As Integer
Dim CombinedHeight As Integer
Dim FOMaterial As clsTrim
Dim TrimPiece As clsTrim
Dim item As Object
Dim tPiece1Length As Double
Dim tPiece2Length As Double
Dim tPiece3Length As Double
Dim tColor As String
Dim m As Integer
Dim jTrimTotalLength As Integer
Dim jTrimRemainder As Integer
Dim jTrim20FtPieces As Integer
Dim Trim20FtPieceCount As Integer
Dim AltGrouping As Boolean
Dim jTrimRemaining As Integer
Dim SplitJambTrim As Boolean
'BPP Solver Adaptation Vars
Dim BPP_TrimCollection As Collection
Dim NumTrimPieces As Integer
''' Debug flag for skipping section
Dim debugFlag As Boolean


''' debug mode
debugFlag = False
If debugFlag = True Then Exit Sub

'find FO Trim color
tColor = EstSht.Range("FO_tColor").Value


''' 'generate jamb trim collection
Set BPP_TrimCollection = New Collection
For Each FO In FOs
    'determine the number of trim pieces
    Select Case FO.Height
    ''' one piece of trim
    Case Is <= ((20 * 12) + 4)
        Set FOMaterial = New clsTrim
        NumTrimPieces = 1
        FOMaterial.tLength = FO.Height
        FOMaterial.tMeasurement = ImperialMeasurementFormat(FO.Height)
        FOMaterial.Quantity = 2
        BPP_TrimCollection.Add FOMaterial
    ''' two pieces of trim
    Case Is <= ((20 * 2 * 12) + (4 * 2) - 2)
        tPiece1Length = NearestTrimSize((FO.Height / 2) + 1, 0, "Jamb", True)
        tPiece2Length = NearestTrimSize((FO.Height - (tPiece1Length - 1)) + 1, 0, "Jamb", True)
        'add directly to material collection for now
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece1Length
        FOMaterial.Quantity = 2
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece1Length)
        FOMaterial.tType = "Jamb Trim"
        BPP_TrimCollection.Add FOMaterial
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece2Length
        FOMaterial.Quantity = 2
        FOMaterial.tType = "Jamb Trim"
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece2Length)
        BPP_TrimCollection.Add FOMaterial
    ''' three pieces of trim
    Case Else
        'find best overlapping trim sizes (accounting for 1" overlap)
        tPiece1Length = NearestTrimSize((FO.Height / 3) + 1, 0, "Jamb", True)
        tPiece2Length = NearestTrimSize((FO.Height / 3) + 1, 0, "Jamb", True)
        'add 2" overlap to the middle piece
        tPiece3Length = NearestTrimSize(FO.Height - (tPiece1Length - 1) - (tPiece2Length - 1) + 2, 0, "Jamb", True)
        'add directly to material collection for now
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece1Length
        FOMaterial.Quantity = 2
        FOMaterial.tType = "Jamb Trim"
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece1Length)
        BPP_TrimCollection.Add FOMaterial
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece2Length
        FOMaterial.tType = "Jamb Trim"
        FOMaterial.Quantity = 2
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece2Length)
        MaterialCollection.Add FOMaterial
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece3Length
        FOMaterial.Quantity = 2
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece3Length)
        BPP_TrimCollection.Add FOMaterial
    End Select
Next FO
'solve jamb trim collection
Call DuplicateMaterialRemoval(BPP_TrimCollection, "Trim")
Call JankyBPPSolver.BPP_Solver(MaterialCollection, BPP_TrimCollection, "Jamb", FOType)
''' 'generate Head trim collection
Set BPP_TrimCollection = New Collection
For Each FO In FOs
    'determine the number of trim pieces
    Select Case FO.Width
    ''' one piece of trim
    Case Is <= ((20 * 12) + 4)
        Set FOMaterial = New clsTrim
        NumTrimPieces = 1
        FOMaterial.tLength = FO.Width
        FOMaterial.tMeasurement = ImperialMeasurementFormat(FO.Width)
        FOMaterial.Quantity = 1
        BPP_TrimCollection.Add FOMaterial
    ''' two pieces of trim
    Case Is <= ((20 * 2 * 12) + (4 * 2) - 2)
        tPiece1Length = NearestTrimSize((FO.Width / 2) + 1, 0, "Head", True)
        tPiece2Length = NearestTrimSize((FO.Width - (tPiece1Length - 1)) + 1, 0, "Head", True)
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece1Length
        FOMaterial.tType = "Head Trim W/ Kickout"
        FOMaterial.Quantity = 1
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece1Length)
        BPP_TrimCollection.Add FOMaterial
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece2Length
        FOMaterial.tType = "Head Trim W/ Kickout"
        FOMaterial.Quantity = 1
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece2Length)
        BPP_TrimCollection.Add FOMaterial
    ''' three pieces of trim
    Case Else
        'find best overlapping trim sizes (accounting for 1" overlap)
        tPiece1Length = NearestTrimSize((FO.Width / 3) + 1, 0, "Head", True)
        tPiece2Length = NearestTrimSize((FO.Width / 3) + 1, 0, "Head", True)
        'add 2" overlap to the middle piece
        tPiece3Length = NearestTrimSize(FO.Width - (tPiece1Length - 1) - (tPiece2Length - 1) + 2, 0, "Head", True)
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece1Length
        FOMaterial.Quantity = 1
        FOMaterial.tType = "Head Trim W/ Kickout"
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece1Length)
        BPP_TrimCollection.Add FOMaterial
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece2Length
        FOMaterial.Quantity = 1
        FOMaterial.tType = "Head Trim W/ Kickout"
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece2Length)
        BPP_TrimCollection.Add FOMaterial
        Set FOMaterial = New clsTrim
        FOMaterial.tLength = tPiece3Length
        FOMaterial.Quantity = 1
        FOMaterial.tType = "Head Trim W/ Kickout"
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece3Length)
        BPP_TrimCollection.Add FOMaterial
    End Select
Next FO
'solve Head Trim collection
Call DuplicateMaterialRemoval(BPP_TrimCollection, "Trim")
Call JankyBPPSolver.BPP_Solver(MaterialCollection, BPP_TrimCollection, "Head", FOType)

'add duplicate head trip without kickout for Windows and Misc Fos
If FOType = "Windows" Or FOType = "MiscFOs" Then
    For Each item In MaterialCollection
    If item.clsType = "Trim" Then
        Set FOMaterial = item
        'add duplicate head trim without kickout
        If FOMaterial.tType = "Head Trim W/ Kickout" Then
            Set TrimPiece = New clsTrim
            TrimPiece.Quantity = FOMaterial.Quantity
            TrimPiece.tMeasurement = FOMaterial.tMeasurement
            TrimPiece.tType = "Head Trim W/O Kickout"
            MaterialCollection.Add TrimPiece
        End If
    End If
    Next item
End If
    
'    'add duplicate head trim without kickout if a Misc FO
'    If FOType = "MiscFOs" Then
'        For m = 1 To MaterialCollection.Count
'            With MaterialCollection(m)
'                If .tType = "Head Trim W/ Kickout" Then
'                    'add equivalent head trim without kickout
'                     Set FOMaterial = New clsTrim
'                    FOMaterial.Color = tColor
'                    FOMaterial.tType = "Head Trim W/O Kickout"
'                    FOMaterial.tMeasurement = .tMeasurement
'                    FOMaterial.tLength = .tLength
'                    FOMaterial.Quantity = 1
'                    MaterialCollection.Add FOMaterial
'                End If
'            End With
'        Next m
'    End If

'add trim color
For Each item In MaterialCollection
If item.clsType = "Trim" Then
    Set FOMaterial = item
    Select Case FOMaterial.tType
    Case "Head Trim W/ Kickout", "Head Trim W/O Kickout", "Jamb Trim"
        FOMaterial.Color = tColor
    Case Else
        FOMaterial.Color = "N/A"
    End Select
End If
Next item

End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sub for calculating panel collections for the roof
'Note: Panel uses Rafter lengths with overhangs already factored in
Private Sub RoofPanelGen(ByRef PanelCollection As Collection, RafterSheetLength As Double, EaveOverhang As Double, RoofLength As Integer, Optional rShape As String, Optional EaveExtension As Boolean)
Dim Overlap As Integer
Dim p1 As clsPanel
Dim p2 As clsPanel
Dim p3 As clsPanel
Dim p4 As clsPanel
Dim p5 As clsPanel

Dim p1Length As Double
Dim p1Measurement As String
Dim p2Length As Double
Dim p2Measurement As String
Dim p3Length As Double
Dim p3Measurement As String
Dim p4Length As Double
Dim p4Measurement As String
Dim p5Length As Double
Dim p5Measurement As String
Dim PanelQty As Integer
Dim IdealPLength As Double
Dim RemainingLength As Double
Dim LargeOverhang As Boolean
Dim OnePanel As Boolean
Dim TwoPanel As Boolean
Dim ThreePanel As Boolean
Dim FourPanel As Boolean
Dim FivePanel As Boolean
Dim SixPanel As Boolean


'standard panel overlap of 6"
Overlap = 6

' Check for overhang greater than 1.5 ft
If EaveOverhang > (1.5 * 12) Then LargeOverhang = True
'''check for next panel size
Select Case True
Case RafterSheetLength <= (42 * 12)
    OnePanel = True
'two panels
Case RafterSheetLength <= (83 * 12)
    If LargeOverhang = False Then
        p1Length = (40 * 12) + EaveOverhang
    Else
        p1Length = (35 * 12) + EaveOverhang
    End If
    p2Length = RafterSheetLength - p1Length
    If p2Length <= (42 * 12) Then
        TwoPanel = True
    Else
        ThreePanel = True
    End If
'three panels
Case RafterSheetLength <= (124 * 12)
    If LargeOverhang = False Then
        p1Length = (40 * 12) + EaveOverhang
    Else
        p1Length = (35 * 12) + EaveOverhang
    End If
    p2Length = (40 * 12)
    p3Length = RafterSheetLength - p1Length - p2Length
    If p3Length <= (42 * 12) Then
        ThreePanel = True
    Else
        FourPanel = True
    End If
'four panels
Case RafterSheetLength <= (165 * 12)
    If LargeOverhang = False Then
        p1Length = (40 * 12) + EaveOverhang
    Else
        p1Length = (35 * 12) + EaveOverhang
    End If
    p2Length = (40 * 12)
    p3Length = p2Length
    p4Length = RafterSheetLength - p1Length - p2Length - p3Length
    If p4Length <= (42 * 12) Then
        FourPanel = True
    Else
        FivePanel = True
    End If
'five panels
Case RafterSheetLength <= (206 * 12)
    If LargeOverhang = False Then
        p1Length = (40 * 12) + EaveOverhang
    Else
        p1Length = (35 * 12) + EaveOverhang
    End If
    p2Length = (40 * 12)
    p3Length = p2Length
    p4Length = p2Length
    p5Length = RafterSheetLength - p1Length - p2Length - p3Length - p4Length
    If p5Length <= (42 * 12) Then
        FivePanel = True
    Else
        SixPanel = True
    End If
Case Else
    SixPanel = True
End Select

' panel quantity
PanelQty = Application.WorksheetFunction.RoundUp((RoofLength / 3), 0)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' roof panels for one sidewall'''''
' at least 1 panel
Set p1 = New clsPanel

'Debug.Print "Rafter Measurement: " & ImperialMeasurementFormat(RafterLength)
'Debug.Print "Rafter Length: " & RafterLength / 12
'Debug.Print "Ideal Panel Length: " & (RafterLength / 12) / 2
'Debug.Print "Ideal Panel Length (minus overhang): " & ((RafterLength - EaveOverhang) / 12) / 2

''' determine how many panels to make
Select Case True
Case OnePanel
    p1.PanelLength = RafterSheetLength
    'add underlap for eave extension panels
    If EaveExtension = True Then p1.PanelLength = p1.PanelLength + 6
    ''' convert to imperial
    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength)
    'quantity
    p1.Quantity = PanelQty
    'add to collection
    PanelCollection.Add p1
    
'' check for 2 divisions
Case TwoPanel
    'new panel class
    Set p2 = New clsPanel
    'include overhang on ideal panel length when comparing 2 panel options
    IdealPLength = RafterSheetLength / 2
    'determine panel 1 length
    If LargeOverhang = False Then
        'round p1 length to panel length that doesn't exceed 40'
        If ClosestRoofPurlin(IdealPLength, 1) <= 40 * 12 Then
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang
        Else
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
        End If
    ElseIf LargeOverhang = True Then
        'round p1 length to panel length that doesn't exceed 35
        If ClosestRoofPurlin(IdealPLength, 1) <= 35 * 12 Then
            'manually check
            If Abs(IdealPLength - (RafterSheetLength - (ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang))) < Abs(IdealPLength - (RafterSheetLength - (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang))) Then
                p1.PanelLength = ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang
            Else
                p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
            End If
        Else
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
        End If
    End If
    'add in overhang
    'p1.PanelLength = p1.PanelLength + EaveOverhang
    'find p2 length
    p2.PanelLength = RafterSheetLength - p1.PanelLength
    'add overlap
    p2.PanelLength = p2.PanelLength + Overlap
    p1.PanelLength = p1.PanelLength + Overlap
    
    'convert lengths, calculate quantities
    ''' convert to imperial
    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength)
    p2.PanelMeasurement = ImperialMeasurementFormat(p2.PanelLength)
    'quantity
    p1.Quantity = PanelQty
    p2.Quantity = PanelQty
    'add to collection
    PanelCollection.Add p1
    PanelCollection.Add p2
    
'' check for 3
Case ThreePanel

    'new panel classes
    Set p2 = New clsPanel
    Set p3 = New clsPanel
    IdealPLength = RafterSheetLength / 3
    
    'determine panel 1 length (short side)
    If LargeOverhang = False Then
        'round p1 length to panel length that doesn't exceed 40'
        If ClosestRoofPurlin(IdealPLength, 1) <= 40 * 12 Then
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
        Else
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
        End If
    ElseIf LargeOverhang = True Then
        'round p1 length to panel length that doesn't exceed 35
        If ClosestRoofPurlin(IdealPLength, 1) <= 35 * 12 Then
            p1.PanelLength = ClosestRoofPurlin(IdealPLength) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
        Else
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
            'check about other panel lengths
            If ClosestRoofPurlin(IdealPLength, 1) <= 40 * 12 Then
                p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
            Else
                p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
            End If
        End If
    End If

    'add in overhang
    'p1.PanelLength = p1.PanelLength + EaveOverhang
    
    'find remaining rafter length
    p3.PanelLength = RafterSheetLength - p1.PanelLength - p2.PanelLength
    
    'add overlap, add undercut back in (because it isn't undercut)
    p3.PanelLength = p3.PanelLength + Overlap
    'add two overlaps, deduct overhang and add undercut back in
    p2.PanelLength = p2.PanelLength + (Overlap * 2)
    
    'add overlap for panel 1
    p1.PanelLength = p1.PanelLength + Overlap
    
    'convert lengths, calculate quantities
    ''' convert to imperial
    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength)
    p2.PanelMeasurement = ImperialMeasurementFormat(p2.PanelLength)
    p3.PanelMeasurement = ImperialMeasurementFormat(p3.PanelLength)
    'quantity
    p1.Quantity = PanelQty
    p2.Quantity = PanelQty
    p3.Quantity = PanelQty
    'add to collection
    PanelCollection.Add p1
    PanelCollection.Add p2
    PanelCollection.Add p3
    
    'remove duplicates
    Call DuplicateMaterialRemoval(PanelCollection, "Panel")
    
''check for 4
Case FourPanel
    'new panel classes
    Set p2 = New clsPanel
    Set p3 = New clsPanel
    Set p4 = New clsPanel
    'ideal
    IdealPLength = RafterSheetLength / 4
    
    'determine panel 1 length (short side)
    If LargeOverhang = False Then
        'round p1 length to panel length that doesn't exceed 40'
        If ClosestRoofPurlin(IdealPLength, 1) <= 40 * 12 Then
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
            p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, p1.PanelLength + p2.PanelLength))
        Else
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
            p3.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
        End If
    ElseIf LargeOverhang = True Then
        'round p1 length to panel length that doesn't exceed 35
        If ClosestRoofPurlin(IdealPLength, 1) <= 35 * 12 Then
            p1.PanelLength = ClosestRoofPurlin(IdealPLength) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
            p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, p1.PanelLength + p2.PanelLength))
        Else
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
            'check about other panel lengths
            If ClosestRoofPurlin(IdealPLength, 1) <= 40 * 12 Then
                p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
                p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, p1.PanelLength + p2.PanelLength))
            Else
                p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
                p3.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
            End If
        End If
    End If
    'add in overhang
    'p1.PanelLength = p1.PanelLength + EaveOverhang
    'determine remaining length
    p4.PanelLength = RafterSheetLength - p1.PanelLength - p2.PanelLength - p3.PanelLength
    'add overlap, add undercut back in (bottom panel)
    p4.PanelLength = p4.PanelLength + Overlap
    'deduct overhang, add overlap for panel 1 (top panel)
    p1.PanelLength = p1.PanelLength + Overlap
    'add two overlaps for panel 2 deduct overhang and add undercut back (because it isn't overhanging or undercut) (middle panel)
    p2.PanelLength = p2.PanelLength + (Overlap * 2)
    'add two overlaps, deduct overhang and add undercut back (because it isn't overhanging or undercut) (middle panel)
    p3.PanelLength = p3.PanelLength + (Overlap * 2)
    
    'convert lengths, calculate quantities
    ''' convert to imperial
    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength)
    p2.PanelMeasurement = ImperialMeasurementFormat(p2.PanelLength)
    p3.PanelMeasurement = ImperialMeasurementFormat(p3.PanelLength)
    p4.PanelMeasurement = ImperialMeasurementFormat(p4.PanelLength)
    'quantity
    p1.Quantity = PanelQty
    p2.Quantity = PanelQty
    p3.Quantity = PanelQty
    p4.Quantity = PanelQty
    'add to collection
    PanelCollection.Add p1
    PanelCollection.Add p2
    PanelCollection.Add p3
    PanelCollection.Add p4
'' check for 5
Case FivePanel
    'new panel classes
    Set p2 = New clsPanel
    Set p3 = New clsPanel
    Set p4 = New clsPanel
    Set p5 = New clsPanel
    'ideal
    IdealPLength = RafterSheetLength / 5
    
    'determine panel 1 length (short side)
    If LargeOverhang = False Then
        'round p1 length to panel length that doesn't exceed 40'
        If ClosestRoofPurlin(IdealPLength, 1) <= 40 * 12 Then
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
            p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, p1.PanelLength + p2.PanelLength))
            p4.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 4, p1.PanelLength + p2.PanelLength + p3.PanelLength))
        Else
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
            p3.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
            p4.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
        End If
    ElseIf LargeOverhang = True Then
        'round p1 length to panel length that doesn't exceed 35
        If ClosestRoofPurlin(IdealPLength, 1) <= 35 * 12 Then
            p1.PanelLength = ClosestRoofPurlin(IdealPLength) + EaveOverhang
            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
            p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, p1.PanelLength + p2.PanelLength))
            p4.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 4, p1.PanelLength + p2.PanelLength + p3.PanelLength))
        Else
            p1.PanelLength = ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang
            'check about other panel lengths
            If ClosestRoofPurlin(IdealPLength, 1) <= 40 * 12 Then
                'check for other panel lengths
                p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength))
                p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, p1.PanelLength + p2.PanelLength))
                p4.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 4, p1.PanelLength + p2.PanelLength + p3.PanelLength))
            Else
                p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
                p3.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
                p4.PanelLength = ClosestRoofPurlin(IdealPLength, -1)
            End If
        End If
    End If
    'add in overhang
    'p1.PanelLength = p1.PanelLength + EaveOverhang
    
    'determine remaining length
    p5.PanelLength = RafterSheetLength - p1.PanelLength - p2.PanelLength - p3.PanelLength - p4.PanelLength
    
    'add overlap, add undercut back in (bottom panel)
    p5.PanelLength = p5.PanelLength + Overlap
    'deduct overhang, add overlap for panel 1 (top panel)
    p1.PanelLength = p1.PanelLength + Overlap
    'add two overlaps for panel 2 deduct overhang and add undercut back (because it isn't overhanging or undercut) (middle panel)
    p2.PanelLength = p2.PanelLength + (Overlap * 2)
    'add two overlaps for panel 3 deduct overhang and add undercut back (because it isn't overhanging or undercut) (middle panel)
    p3.PanelLength = p3.PanelLength + (Overlap * 2)
     'add two overlaps, deduct overhang and add undercut back (because it isn't overhanging or undercut) (middle panel)
    p4.PanelLength = p4.PanelLength + (Overlap * 2)
    
    
    'convert lengths, calculate quantities
    ''' convert to imperial
    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength)
    p2.PanelMeasurement = ImperialMeasurementFormat(p2.PanelLength)
    p3.PanelMeasurement = ImperialMeasurementFormat(p3.PanelLength)
    p4.PanelMeasurement = ImperialMeasurementFormat(p4.PanelLength)
    p5.PanelMeasurement = ImperialMeasurementFormat(p5.PanelLength)
    'quantity
    p1.Quantity = PanelQty
    p2.Quantity = PanelQty
    p3.Quantity = PanelQty
    p4.Quantity = PanelQty
    p5.Quantity = PanelQty
    'add to collection
    PanelCollection.Add p1
    PanelCollection.Add p2
    PanelCollection.Add p3
    PanelCollection.Add p4
    PanelCollection.Add p5
Case Else
    GoTo LargePanelDivision
End Select


Exit Sub

LargePanelDivision:
MsgBox "It has been calculated that more than 5 seperate panels will be needed to cover the rafter length of the roof. Please perform this calculation manually.", vbExclamation, "Program Scope Exceeded"
End
End Sub


Private Sub SoffitGen(SoffitPanels As Collection, SoffitTrim As Collection, SoffitLocation As String, b As clsBuilding, Optional s2RoofPanels As Collection, Optional s4RoofPanels As Collection)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sub for Generating Soffit Panels and Soffit Trim
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RoofPanel As clsPanel
Dim SoffitPanel As clsPanel
Dim NamedRangeString As String  'Var for reading correct soffit panel/trim info cell
Dim SoffitQty As Integer
Dim NetRafterLength As Double
Dim NetOutsideAngleLength As Double
Dim TrimPiece As clsTrim
Dim SoffitLength As Integer
Dim NetStandardEaveOverhang As Double   'var for subtracting the 4.25" overhang as needed for a single slope
Dim EaveExtBuildingLength As Integer 'eave extension length from endwall to endwall
Dim EaveExtRafterLength As Double

NamedRangeString = SoffitLocation & "Soffit"
With b
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Soffit Panels
    'determine Soffit Panel Quantity
    Select Case SoffitLocation
    Case "e1_GableOverhang"
        SoffitQty = Application.WorksheetFunction.RoundUp(((.e1Overhang / 12) + (.e1ExtensionOverhang / 12)) / 5, 0)
        SoffitLength = .e1Overhang + .e1ExtensionOverhang
    Case "e3_GableOverhang"
        SoffitQty = Application.WorksheetFunction.RoundUp(((.e3Overhang / 12) + (.e3ExtensionOverhang / 12)) / 5, 0)
        SoffitLength = .e3Overhang + .e3ExtensionOverhang
    Case "e1_GableExtension"
        SoffitQty = Application.WorksheetFunction.RoundUp((.e1Extension / 12) / 5, 0)
        SoffitLength = .e1Extension
    Case "e3_GableExtension"
        SoffitQty = Application.WorksheetFunction.RoundUp((.e3Extension / 12) / 5, 0)
        SoffitLength = .e3Extension
    Case "s2_EaveOverhang", "s4_EaveOverhang"
        'eave overhang soffits grouped along just the sidewall length
        SoffitQty = Application.WorksheetFunction.RoundUp(.bLength / 3, 0)
    Case "s2_EaveExtension"
        'eave extension soffits are along entire roof length
        EaveExtBuildingLength = .s2EaveExtensionBuildingLength / 12
        EaveExtRafterLength = .s2ExtensionRafterLength
    Case "s4_EaveExtension"
        EaveExtBuildingLength = .s4EaveExtensionBuildingLength / 12
        EaveExtRafterLength = .s4ExtensionRafterLength
    End Select
    
    'Generate of Gable Overhang/Extension Soffit Panels
    If InStr(1, SoffitLocation, "Gable") <> 0 Then
        If .rShape = "Single Slope" Then
            'subtract the standard eave overhang
            If .s4Overhang <> 0 Or .s4ExtensionOverhang <> 0 Then
                NetStandardEaveOverhang = 4.25 * 2
            Else
                NetStandardEaveOverhang = 4.25
            End If
            'add soffit corresponding to roof panels of sidewall 2, less the standard overhangs
            Call RoofPanelGen(SoffitPanels, .s2RafterSheetLength - NetStandardEaveOverhang, 0, SoffitLength / 12)
            For Each SoffitPanel In SoffitPanels
                SoffitPanel.PanelShape = EstSht.Range(NamedRangeString).offset(0, 1).Value
                SoffitPanel.PanelType = EstSht.Range(NamedRangeString).offset(0, 2).Value
                SoffitPanel.PanelColor = EstSht.Range(NamedRangeString).offset(0, 3).Value
                SoffitPanel.clsType = "Panel"
            Next SoffitPanel
        
        'if a gable roof, don't subtract the eave overhang (due to the undercut) and just match each sidewall's rafter sheet length
        
        ''''''''''' NOTE: DOES THIS NEED TO BE FIXED? PERHAPS NOT...
        ElseIf .rShape = "Gable" Then
            Call RoofPanelGen(SoffitPanels, .s2RafterSheetLength, .s2Overhang, SoffitLength / 12)
            Call RoofPanelGen(SoffitPanels, .s4RafterSheetLength, .s4Overhang, SoffitLength / 12)
            For Each SoffitPanel In SoffitPanels
                SoffitPanel.PanelShape = EstSht.Range(NamedRangeString).offset(0, 1).Value
                SoffitPanel.PanelType = EstSht.Range(NamedRangeString).offset(0, 2).Value
                SoffitPanel.PanelColor = EstSht.Range(NamedRangeString).offset(0, 3).Value
                SoffitPanel.clsType = "Panel"
            Next SoffitPanel
        End If
        
        'Remove Duplicate Soffit Panels
        Call DuplicateMaterialRemoval(SoffitPanels, "Panel")
        
    'Generation of Eave Overhang Soffit Panels
    ElseIf InStr(1, SoffitLocation, "EaveOverhang") <> 0 Then
        Set SoffitPanel = New clsPanel
        'make soffit panel collection
        If SoffitLocation = "s2_EaveOverhang" Then
            SoffitPanel.PanelLength = .s2Overhang + .s2ExtensionOverhangRafterLength
        ElseIf SoffitLocation = "s4_EaveOverhang" Then
            SoffitPanel.PanelLength = .s4Overhang + .s4ExtensionOverhangRafterLength
        End If
        SoffitPanel.PanelMeasurement = ImperialMeasurementFormat(SoffitPanel.PanelLength)
        SoffitPanel.PanelShape = EstSht.Range(NamedRangeString).offset(0, 1).Value
        SoffitPanel.PanelType = EstSht.Range(NamedRangeString).offset(0, 2).Value
        SoffitPanel.PanelColor = EstSht.Range(NamedRangeString).offset(0, 3).Value
        SoffitPanel.Quantity = SoffitQty
        SoffitPanel.clsType = "Panel"
        SoffitPanels.Add SoffitPanel

    'Generation of Eave Extension Soffit Panels
    ElseIf InStr(1, SoffitLocation, "EaveExtension") <> 0 Then
        Call RoofPanelGen(SoffitPanels, EaveExtRafterLength, 0, EaveExtBuildingLength)
        'update panel parameters
        For Each SoffitPanel In SoffitPanels
            SoffitPanel.PanelShape = EstSht.Range(NamedRangeString).offset(0, 1).Value
            SoffitPanel.PanelType = EstSht.Range(NamedRangeString).offset(0, 2).Value
            SoffitPanel.PanelColor = EstSht.Range(NamedRangeString).offset(0, 3).Value
            SoffitPanel.clsType = "Panel"
        Next SoffitPanel
    End If
    
    
    ''''''''''' NOTE: THIS NEEDS TO BE UPDATED
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Soffit Trim
    ''''''''''''''''''''''''''''''''' Gable Overhangs/Extensions
    'Jamb Trim and 2x6 Outside Angle
    If InStr(1, SoffitLocation, "Gable") <> 0 Then
        'calculate net rafter length for one endwall
        If .rShape = "Gable" Then
            NetRafterLength = (.RafterLength * 2)
        ElseIf .rShape = "Single Slope" Then
            NetRafterLength = .RafterLength
        End If
        'generate Jamb Trim unless it's an extension overhang
        If (InStr(1, SoffitLocation, "e1_GableOverhang") <> 0 And .e1ExtensionOverhang > 0) Or (InStr(1, SoffitLocation, "e3_GableOverhang") <> 0 And .e3ExtensionOverhang > 0) Then
            'dont generate jamb trim
        Else
            'normal case jamb trim generation
            Call TrimPieceCalc(SoffitTrim, NetRafterLength, "Jamb", , , b)
        End If
        'Generate 2x6 outside angle trim unless it's an extension with an overhang
        If (InStr(1, SoffitLocation, "e1_GableExtension") <> 0 And .e1ExtensionOverhang > 0) Or (InStr(1, SoffitLocation, "e3_GableExtension") <> 0 And .e3ExtensionOverhang > 0) Then
        ' dont generate outside angle trim
        Else
            'normal case 2x6 outside angle generation
            If .rShape = "Single Slope" Then
                NetOutsideAngleLength = .s2RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength + .s2ExtensionOverhangRafterLength + .s4ExtensionOverhangRafterLength
                Call TrimPieceCalc(SoffitTrim, NetOutsideAngleLength, "Outside Angle", , , b)
            ElseIf .rShape = "Gable" Then
                NetOutsideAngleLength = .s2RafterSheetLength + .s4RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength + .s2ExtensionOverhangRafterLength + .s4ExtensionOverhangRafterLength
                Call TrimPieceCalc(SoffitTrim, NetOutsideAngleLength, "Outside Angle", , , b)
            End If
        End If
        'Update Trim Information
        For Each TrimPiece In SoffitTrim
            With TrimPiece
                If .tType = "Jamb Trim" Then .tShape = b.wPanelShape
                If .tType = "2x6 Outside Angle" Then .tShape = "N/A"
            End With
        Next TrimPiece
    ''''''''''''''''''''''''''''''''' Eave Overhangs/Extensions
    'Head Trim and 2x6 Outside Angle
    ElseIf InStr(1, SoffitLocation, "Eave") <> 0 Then
        'generate head trim unless it's an extension overhang
        If (InStr(1, SoffitLocation, "s2_EaveOverhang") <> 0 And .s2ExtensionOverhang) Or (InStr(1, SoffitLocation, "s4_EaveOverhang") <> 0 And .s4ExtensionOverhang > 0) Then
            'dont generate head trim
        Else
            'normal case - Head trim generation for one sidewall
            Call TrimPieceCalc(SoffitTrim, .bLength * 12, "Head")
        End If
        
        '''''Generate 2x6 outside angle trim
        'Eave Overhang (where 2x6 outside angle def covers the gable extensions/overhangs)
        If InStr(1, SoffitLocation, "Overhang") <> 0 Then
            NetOutsideAngleLength = (.bLength * 12) + .e1Overhang + .e1Extension + .e3Overhang + .e3Extension + .e1ExtensionOverhang + .e3ExtensionOverhang
            Call TrimPieceCalc(SoffitTrim, NetOutsideAngleLength, "Outside Angle", , , b)
        'Eave Extension (where 2x6 outside angle may not cover the gable extensions/overhangs
        ElseIf InStr(1, SoffitLocation, "Extension") <> 0 Then
            'only generate if we're NOT handling an extension with an overhang:
            If (InStr(1, SoffitLocation, "s2") <> 0 And .s2ExtensionOverhang = 0) Or (InStr(1, SoffitLocation, "s4") <> 0 And .s4ExtensionOverhang = 0) Then
                NetOutsideAngleLength = EaveExtBuildingLength * 12
                Call TrimPieceCalc(SoffitTrim, NetOutsideAngleLength, "Outside Angle", , , b)
            End If
        End If
        
        'Update Trim Information
        For Each TrimPiece In SoffitTrim
            With TrimPiece
                If .tType = "Head Trim W/O Kickout" Then .tShape = b.wPanelShape
                If .tType = "2x6 Outside Angle" Then
                    .tShape = "N/A"
                    ''' Roof pitch string for outside angle trim on eave overhangs/extensions
                    'roof pitch may vary on an eave extension
                    If InStr(1, SoffitLocation, "Extension") <> 0 Then
                        If SoffitLocation = "s2_EaveExtension" Then .tType = .tType & " " & b.s2ExtensionPitch & ":12"
                        If SoffitLocation = "s4_EaveExtension" Then .tType = .tType & " " & b.s4ExtensionPitch & ":12"
                    'roof pitch will not vary on an eave overhang, except when in an overhang extension
                    ElseIf InStr(1, SoffitLocation, "Overhang") <> 0 Then
                        If InStr(1, SoffitLocation, "s2") <> 0 And b.s2ExtensionOverhang <> 0 Then   'check if handling an extension overhang on s2
                            .tType = .tType & " " & b.s2ExtensionPitch & ":12"
                        ElseIf InStr(1, SoffitLocation, "s4") <> 0 And b.s4ExtensionOverhang <> 0 Then   'check if handling an extension overhang on s4
                            .tType = .tType & " " & b.s4ExtensionPitch & ":12"
                        Else    'handling normal eave overhangs
                            .tType = .tType & " " & b.rPitch & ":12"
                        End If
                        
                    End If
                End If
            End With
        Next TrimPiece
    End If
End With

End Sub

Private Sub ExtensionPanelGen(ExtensionPanels As Collection, b As clsBuilding, ExtensionLocation As String, Optional s2RoofPanels As Collection, Optional s4RoofPanels As Collection)
Dim RoofPanel As clsPanel
Dim ExtensionPanel As clsPanel
Dim NamedRangeString As String 'Var for reading correct extension panel info cell
Dim PanelQty As Integer
Dim EaveExtBuildingLength As Integer 'length measured from endwall to endwall
Dim EaveExtRafterLength As Double
Dim ExtensionLengthOverage As Integer
NamedRangeString = ExtensionLocation

With b
    Select Case ExtensionLocation
    Case "e1_GableExtension"
        'check for partial extension panel usage
        ExtensionLengthOverage = (.e1Extension + .e1ExtensionOverhang) Mod (3 * 12)
        If ExtensionLengthOverage > 0 And .bLengthRoofPanelOverage >= ExtensionLengthOverage Then
            'use roof panel overage
            PanelQty = Application.WorksheetFunction.RoundUp(((.e1Extension + .e1ExtensionOverhang - ExtensionLengthOverage) / 12) / 3, 0)
            'update roof panel overage remaining
            .bLengthRoofPanelOverage = .bLengthRoofPanelOverage - ExtensionLengthOverage
        Else
            PanelQty = Application.WorksheetFunction.RoundUp(((.e1Extension + .e1ExtensionOverhang) / 12) / 3, 0)
        End If
        'set extension panel quantity
        .e1ExtensionPanelQty = PanelQty
    Case "e3_GableExtension"
        'check for partial extension panel usage
        ExtensionLengthOverage = (.e3Extension + .e3ExtensionOverhang) Mod (3 * 12)
        If ExtensionLengthOverage > 0 And .bLengthRoofPanelOverage >= ExtensionLengthOverage Then
            'use roof panel overage
            PanelQty = Application.WorksheetFunction.RoundUp(((.e3Extension + .e3ExtensionOverhang - ExtensionLengthOverage) / 12) / 3, 0)
            'update roof panel overage remaining
            .bLengthRoofPanelOverage = .bLengthRoofPanelOverage - ExtensionLengthOverage
        Else
            PanelQty = Application.WorksheetFunction.RoundUp(((.e3Extension + .e3ExtensionOverhang) / 12) / 3, 0)
        End If
        .e3ExtensionPanelQty = PanelQty
    Case "s2_EaveExtension"
        'length from endwall to endwall
        EaveExtBuildingLength = .s2EaveExtensionBuildingLength / 12
        'eave ext rafter length
        EaveExtRafterLength = .s2ExtensionRafterLength + .s2ExtensionOverhangRafterLength
    Case "s4_EaveExtension"
        EaveExtBuildingLength = .s4EaveExtensionBuildingLength / 12
        'eave ext rafter length
        EaveExtRafterLength = .s4ExtensionRafterLength + .s4ExtensionOverhangRafterLength
    End Select
    
    ''''''''''''''''''''''''''''''''''''''''''''''' For Gable Extensions '''''''''''''''''''''''''''''''''''''''''
    If InStr(1, ExtensionLocation, "GableExtension") <> 0 Then
        ''' Corresponding Extension Panels for Each Sidewall 2 Roof Panel type
        For Each RoofPanel In s2RoofPanels
            Set ExtensionPanel = New clsPanel
            ExtensionPanel.PanelMeasurement = RoofPanel.PanelMeasurement
            ExtensionPanel.PanelShape = .rPanelShape
            ExtensionPanel.PanelType = .rPanelType
            ExtensionPanel.PanelColor = .rPanelColor
            ExtensionPanel.Quantity = PanelQty
            ExtensionPanel.clsType = "Panel"
            ExtensionPanels.Add ExtensionPanel
        Next RoofPanel
            
        'For a Gable Roof, Corresponding Extension Panels for Each Sidewall 4 Roof Panel type
        If .rShape = "Gable" Then
            For Each RoofPanel In s4RoofPanels
                Set ExtensionPanel = New clsPanel
                ExtensionPanel.PanelMeasurement = RoofPanel.PanelMeasurement
                ExtensionPanel.PanelShape = .rPanelShape
                ExtensionPanel.PanelType = .rPanelType
                ExtensionPanel.PanelColor = .rPanelColor
                ExtensionPanel.Quantity = PanelQty
                ExtensionPanel.clsType = "Panel"
                ExtensionPanels.Add ExtensionPanel
            Next RoofPanel
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Eave Extension Panels
    ElseIf InStr(1, ExtensionLocation, "EaveExtension") <> 0 Then
        If InStr(1, ExtensionLocation, "s2") <> 0 Then
            Call RoofPanelGen(ExtensionPanels, EaveExtRafterLength, .s2ExtensionOverhangRafterLength, EaveExtBuildingLength, b.rShape, True)
        ElseIf InStr(1, ExtensionLocation, "s4") <> 0 Then
            Call RoofPanelGen(ExtensionPanels, EaveExtRafterLength, .s4ExtensionOverhangRafterLength, EaveExtBuildingLength, b.rShape, True)
        End If

        'update panel parameters
        For Each ExtensionPanel In ExtensionPanels
            ExtensionPanel.PanelShape = .rPanelShape
            ExtensionPanel.PanelType = .rPanelType
            ExtensionPanel.PanelColor = .rPanelColor
            ExtensionPanel.clsType = "Panel"
        Next ExtensionPanel
    End If
    
     'remove duplicate panels
    Call DuplicateMaterialRemoval(ExtensionPanels, "Panel")
End With
        

End Sub

Private Function PanelOptionCompare(IdealPLength As Double, PanelCount As Integer, CurrentTotalLength As Double) As Integer
''' Function to determine whether or not to round the roof panel up or down to the nearest purlin to keep the total closest to the ideal
If Abs((IdealPLength * PanelCount) - (CurrentTotalLength + ClosestRoofPurlin(IdealPLength, 1))) < Abs((IdealPLength * PanelCount) - (CurrentTotalLength + ClosestRoofPurlin(IdealPLength, -1))) Then
    PanelOptionCompare = 1
Else
    PanelOptionCompare = -1
End If

End Function

Private Sub RoofScrewGen(TekQty As Integer, LapQty As Integer, b As clsBuilding, rOverlaps As Integer)
Dim xLapSpaces As Integer
Dim yLapSpaces As Integer
Dim rPurlins As Integer
Dim s2ExtensionPurlins As Integer
Dim s4ExtensionPurlins As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Roof Screws ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
With b
    rPurlins = Application.WorksheetFunction.RoundUp((.RafterLength / 12) / 5, 0)
    'double purlins for gable roof
    If .rShape = "Gable" Then rPurlins = rPurlins * 2
    '''add in overhang/extensions
    'sidewall 2 roof purlins
    If .s2Extension = 0 Then
        'add in one purlin for an additional overhang
        If .s2Overhang > 4.25 Then rPurlins = rPurlins + 1
    ElseIf .s2Extension <> 0 Then
        s2ExtensionPurlins = Application.WorksheetFunction.RoundUp(((.s2ExtensionRafterLength + .s2ExtensionOverhangRafterLength) / 12) / 5, 0)
        rPurlins = rPurlins + s2ExtensionPurlins
    End If
    'sidewall 4 roof purlins
    If .s4Extension = 0 Then
        'add in one purlin for an eave extension
        If .s4Overhang > 4.25 Then rPurlins = rPurlins + 1
    ElseIf .s4Extension <> 0 Then
        s4ExtensionPurlins = Application.WorksheetFunction.RoundUp(((.s4ExtensionRafterLength + .s4ExtensionOverhangRafterLength) / 12) / 5, 0)
        rPurlins = rPurlins + s4ExtensionPurlins
    End If
    
                
    ''''calculate tek screw quantity
    'one for every purlin per foot length, top and bottom putlin get 2 per ft length, and overlaps get two per ft length
    If .rShape = "Single Slope" Then
        TekQty = (rPurlins * .RoofFtLength) + (2 * .RoofFtLength) + (rOverlaps * .RoofFtLength)
    ElseIf .rShape = "Gable" Then
        'additional top and bottom purlin for s4
        TekQty = (rPurlins * .RoofFtLength) + (4 * .RoofFtLength) + (rOverlaps * .RoofFtLength)
    End If
    'exclude intersections
    'for sidewall 2 intersections
    If .s2Extension > 0 Then
        If .e1Extension > 0 And .s2e1ExtensionIntersection = False Then TekQty = TekQty - (s2ExtensionPurlins * (.e1Extension / 12))
        If .e3Extension > 0 And .s2e3ExtensionIntersection = False Then TekQty = TekQty - (s2ExtensionPurlins * (.e3Extension / 12))
    End If
    'for sidewall 4 intersections
    If .s4Extension > 0 Then
        If .e1Extension > 0 And .s4e1ExtensionIntersection = False Then TekQty = TekQty - (s4ExtensionPurlins * (.e1Extension / 12))
        If .e3Extension > 0 And .s4e3ExtensionIntersection = False Then TekQty = TekQty - (s4ExtensionPurlins * (.e3Extension / 12))
    End If
    'round up tek screw to the nearest 250
    TekQty = Application.WorksheetFunction.RoundUp(TekQty / 250, 0) * 250
    
    
    '''''calculate lap screw quantity
    'roof length spacses
    yLapSpaces = Application.WorksheetFunction.RoundUp((.RoofLength / 12) / 3, 0) + 1
    If .rShape = "Single Slope" Then
        'rafter length /3'
        xLapSpaces = Application.WorksheetFunction.RoundUp((.s2RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength + .s2ExtensionOverhangRafterLength + .s4ExtensionOverhangRafterLength) / 30, 0) + 1
        'calculate lap qty
        LapQty = (xLapSpaces * yLapSpaces)
    ElseIf .rShape = "Gable" Then
        'sidewall 2
        xLapSpaces = Application.WorksheetFunction.RoundUp(((.s2RafterSheetLength + .s2ExtensionRafterLength + .s2ExtensionOverhangRafterLength)) / 30, 0) + 1
        LapQty = (xLapSpaces * yLapSpaces)
        'sidewall 4
        xLapSpaces = Application.WorksheetFunction.RoundUp((.s4RafterSheetLength + .s4ExtensionRafterLength + .s4ExtensionOverhangRafterLength) / 30, 0) + 1
        LapQty = LapQty + (xLapSpaces * yLapSpaces)
    End If
    
    ''''''''''''''''''''''''''' Note: Lap screws NOT** Currently reduced for excluded extension intersections
    
    'increase if gutters
    If .Gutters = True Then
        'add additional screws for gutters along s2
        LapQty = LapQty + (2 * .RoofFtLength)
        'if gable, add more screws for gutters on s4
        If .rShape = "Gable" Then LapQty = LapQty + (2 * .RoofFtLength)
    End If
    
    'round up
    LapQty = Application.WorksheetFunction.RoundUp(LapQty / 250, 0) * 250
    
End With

End Sub

Private Sub WallScrewGen(TekQty As Integer, LapQty As Integer, b As clsBuilding, sOverlaps As Integer, eOverlaps As Integer)
Dim sPurlins As Integer
Dim ePurlins As Integer
Dim sTekScrews As Integer
Dim eTekScrews As Integer
Dim RemainingHeight As Double
Dim xLapSpaces As Integer
Dim yLapSpaces As Integer
Dim e1HeightIncriment As Double
Dim e3HeightIncriment As Double
Dim s2RemainingHeight As Double
Dim s4RemainingHeight As Double
Dim e1Purlins As Integer
Dim e3Purlins As Integer
Dim s2Purlins As Integer
Dim s4Purlins As Integer
Dim e1MaxLength As Double
Dim e3MaxLength As Double
Dim s2MaxLength As Double
Dim s4MaxLength As Double
Dim e1GablePurlinTotalLength As Double
Dim e3GablePurlinTotalLength As Double
Dim e1WallPurlinTotalLength As Double
Dim e3WallPurlinTotalLength As Double
Dim TotalGableTekScrews As Integer
Dim InteriorRoofAngle As Variant
Dim WallHeightCounter As Integer

With b
    'calculate interior roof angle
    InteriorRoofAngle = WorksheetFunction.Asin(.rPitch / (Sqr(.rPitch ^ 2 + 12 ^ 2)))
    '''''''''''''''''''''''''''''''' Find Purlin Count
    '''calculate endwall purlins
    'peak height
    If .rShape = "Single Slope" Then
        Select Case b.WallStatus("e1")
        Case "Exclude"
            e1MaxLength = 0
        Case "Include", "Partial"
            e1MaxLength = (.bHeight - b.LengthAboveFinishedFloor("e1")) + ((.rPitch * .bWidth) / 12)
        Case "Gable Only"
            e1MaxLength = ((.rPitch * .bWidth) / 12)
        End Select
        Select Case b.WallStatus("e3")
        Case "Exclude"
            e3MaxLength = 0
        Case "Include", "Partial"
            e3MaxLength = (.bHeight - b.LengthAboveFinishedFloor("e3")) + ((.rPitch * .bWidth) / 12)
        Case "Gable Only"
            e3MaxLength = ((.rPitch * .bWidth) / 12)
        End Select
    ElseIf .rShape = "Gable" Then
        Select Case b.WallStatus("e1")
        Case "Exclude"
            e1MaxLength = 0
        Case "Include", "Partial"
            e1MaxLength = (.bHeight - b.LengthAboveFinishedFloor("e1")) + ((.rPitch * (.bWidth / 2)) / 12)
        Case "Gable Only"
            e1MaxLength = ((.rPitch * (.bWidth / 2)) / 12)
        End Select
        Select Case b.WallStatus("e3")
        Case "Exclude"
            e3MaxLength = 0
        Case "Include", "Partial"
            e3MaxLength = (.bHeight - b.LengthAboveFinishedFloor("e3")) + ((.rPitch * (.bWidth / 2)) / 12)
        Case "Gable Only"
            e3MaxLength = ((.rPitch * (.bWidth / 2)) / 12)
        End Select
    End If
    'account for bottom purlins
    If b.WallStatus("e1") = "Include" Then
        e1HeightIncriment = 7 + (2 / 12)
    ElseIf b.WallStatus("e1") <> "Exclude" Then
        e1HeightIncriment = 5
    End If
    If b.WallStatus("e3") = "Include" Then
        e3HeightIncriment = 7 + (2 / 12)
    ElseIf b.WallStatus("e3") <> "Exclude" Then
        e3HeightIncriment = 5
    End If
    
    If .WallStatus("e1") <> "Exclude" Then
        Do
            e1Purlins = e1Purlins + 1
            Select Case b.WallStatus("e1")
            Case "Include", "Partial"
                If e1HeightIncriment > (.bHeight - b.LengthAboveFinishedFloor("e1")) Then
                    e1GablePurlinTotalLength = e1GablePurlinTotalLength + ((e1MaxLength - e1HeightIncriment) / Tan(InteriorRoofAngle))
                Else
                    e1WallPurlinTotalLength = e1WallPurlinTotalLength + .bWidth
                End If
            Case "Gable"
                e1GablePurlinTotalLength = e1GablePurlinTotalLength + ((e1MaxLength - e1HeightIncriment) / Tan(InteriorRoofAngle))
            End Select
            e1HeightIncriment = e1HeightIncriment + 5
        Loop While e1HeightIncriment < e1MaxLength
    End If
    
    If .WallStatus("e3") <> "Exclude" Then
         Do
            e3Purlins = e3Purlins + 1
            Select Case b.WallStatus("e3")
            Case "Include", "Partial"
                If e3HeightIncriment > (.bHeight - b.LengthAboveFinishedFloor("e3")) Then
                    e3GablePurlinTotalLength = e3GablePurlinTotalLength + ((e3MaxLength - e3HeightIncriment) / Tan(InteriorRoofAngle))
                Else
                    e3WallPurlinTotalLength = e3WallPurlinTotalLength + .bWidth
                End If
            Case "Gable"
                e3GablePurlinTotalLength = e3GablePurlinTotalLength + ((e3MaxLength - e3HeightIncriment) / Tan(InteriorRoofAngle))
            End Select
            e3HeightIncriment = e3HeightIncriment + 5
        Loop While e3HeightIncriment < e3MaxLength
    End If
    

    If .rShape = "Gable" Then
        e3GablePurlinTotalLength = e3GablePurlinTotalLength * 2
        e1GablePurlinTotalLength = e1GablePurlinTotalLength * 2
    End If
    If .WallStatus("e1") = "Gable Only" Then
        e1GablePurlinTotalLength = e1GablePurlinTotalLength + .bWidth
    ElseIf .WallStatus("e1") <> "Exclude" Then
        e1WallPurlinTotalLength = e1WallPurlinTotalLength + .bWidth
    End If
    If .WallStatus("e3") = "Gable Only" Then
        e3GablePurlinTotalLength = e3GablePurlinTotalLength + .bWidth
    ElseIf .WallStatus("e1") <> "Exclude" Then
        e1WallPurlinTotalLength = e1WallPurlinTotalLength + .bWidth
    End If
    'purlin length of top not accounted for?
    '''''''''''''calculate sidewall purlins
    Select Case b.WallStatus("s2")
    Case "Exclude"
        s2MaxLength = 0
    Case "Include", "Partial"
        s2MaxLength = .bHeight - b.LengthAboveFinishedFloor("s2")
    End Select
    Select Case b.WallStatus("s4")
    Case "Exclude"
        s4MaxLength = 0
    Case "Include", "Partial"
        If .rShape = "Gable" Then
            s4MaxLength = .bHeight - b.LengthAboveFinishedFloor("s4")
        ElseIf .rShape = "Single Slope" Then
            s4MaxLength = (.HighSideEaveHeight / 12) - b.LengthAboveFinishedFloor("s4")
        End If
    End Select
    'account for bottom purlins
    If b.WallStatus("s2") = "Include" Then
        s2RemainingHeight = s2MaxLength - 7 - (2 / 12)
    ElseIf b.WallStatus("s2") <> "Exclude" Then
        s2RemainingHeight = s2MaxLength - 5
    End If
    If b.WallStatus("s4") = "Include" Then
        s4RemainingHeight = s4MaxLength - 7 - (2 / 12)
    ElseIf b.WallStatus("s4") <> "Exclude" Then
        s4RemainingHeight = s4MaxLength - 5
    End If
    sPurlins = 1
    'find purlins above
    Do While s2RemainingHeight >= 5
        s2Purlins = s2Purlins + 1
        s2RemainingHeight = s2RemainingHeight - 5
    Loop
    Do While s4RemainingHeight >= 5
        s4Purlins = s4Purlins + 1
        s4RemainingHeight = s4RemainingHeight - 5
    Loop
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Tek screws
    ' one per ft per purlin, two for top and bottom purlin
    sTekScrews = (s2Purlins * .bLength) + (2 * .bLength) + (s4Purlins * .bLength) + (2 * .bLength) + (sOverlaps * .bLength)
    'endwall tek screws
    'screws for normal bheight, two for top and bottom purlins
    'eTekScrews = (e1Purlins * .bWidth) + (2 * .bWidth) + (e3Purlins * .bWidth) + (2 * .bWidth) + (eOverlaps * .bWidth)
    eTekScrews = e1WallPurlinTotalLength + e1GablePurlinTotalLength + e3WallPurlinTotalLength + e3GablePurlinTotalLength + (eOverlaps * .bWidth)
    'add together, round up
    TekQty = Application.WorksheetFunction.RoundUp((sTekScrews + eTekScrews) / 250, 0) * 250
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Lap Screws
    ''sidewall 2
    xLapSpaces = Application.WorksheetFunction.RoundUp(.bLength / 3, 0) + 1
    yLapSpaces = Application.WorksheetFunction.RoundUp((s2MaxLength * 12) / 30, 0) + 1
    LapQty = (xLapSpaces * yLapSpaces)
    ''sidewall 4
    yLapSpaces = Application.WorksheetFunction.RoundUp((s4MaxLength * 12) / 30, 0) + 1
    LapQty = LapQty + (xLapSpaces * yLapSpaces)
    '''endwalls
    'square around endwall 1 and 3
    xLapSpaces = Application.WorksheetFunction.RoundUp(.bWidth / 3, 0) + 1
    yLapSpaces = Application.WorksheetFunction.RoundUp((e1MaxLength * 12) / 30, 0) + 1
    LapQty = (xLapSpaces * yLapSpaces)
    yLapSpaces = Application.WorksheetFunction.RoundUp((e3MaxLength * 12) / 30, 0) + 1
    LapQty = LapQty + (xLapSpaces + yLapSpaces)
'    'add extra pitch due to roof (a square of the max building height and the building width
'    If .rShape = "Single Slope" Then
'        yLapSpaces = Application.WorksheetFunction.RoundUp(((.bHeight * 12) + (.rPitch * .bWidth)) / 30, 0) + 1
'    ElseIf .rShape = "Gable" Then
'        yLapSpaces = Application.WorksheetFunction.RoundUp(((.bHeight * 12) + (.rPitch * (.bWidth / 2))) / 30, 0) + 1
'    End If
'    LapQty = LapQty + (xLapSpaces * yLapSpaces)
'
    'round up
    LapQty = Application.WorksheetFunction.RoundUp(LapQty / 250, 0) * 250
    
End With
End Sub

Private Sub TrimScrewCalc(TrimScrews As Collection, RakeTrimPieces As Collection, b As clsBuilding)
Dim NetCornerLength As Integer
Dim NetRakeTrimLength As Integer
Dim Screw As clsFastener
Dim TrimPiece As clsTrim


'calculate net rake trim length
For Each TrimPiece In RakeTrimPieces
    NetRakeTrimLength = NetRakeTrimLength + (TrimPiece.tLength * TrimPiece.Quantity)
Next TrimPiece

'calculate corner trim length
If b.rShape = "Gable" Then
    'assume complete, exclude intersections if needed
    NetCornerLength = b.bHeight * 4 * 12
    If b.WallStatus("e1") <> "Include" And b.WallStatus("s2") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("e1") <> "Include" And b.WallStatus("s4") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("e3") <> "Include" And b.WallStatus("s2") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("e3") <> "Include" And b.WallStatus("s4") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
ElseIf b.rShape = "Single Slope" Then
    'sidewall 2 corners + s4 corners
    NetCornerLength = (b.bHeight * 12 * 2) + (b.HighSideEaveHeight * 2)
    'exclude where needed
    If b.WallStatus("s2") <> "Include" And b.WallStatus("e1") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("s2") <> "Include" And b.WallStatus("e3") <> "Include" Then NetCornerLength = NetCornerLength - (b.bHeight * 12)
    If b.WallStatus("s4") <> "Include" And b.WallStatus("e1") <> "Include" Then NetCornerLength = NetCornerLength - b.HighSideEaveHeight
    If b.WallStatus("s4") <> "Include" And b.WallStatus("e3") <> "Include" Then NetCornerLength = NetCornerLength - b.HighSideEaveHeight
End If

''' Add to Screws Collection
' If screws the same color, combine
If b.OutsideCornerTrimColor = b.RakeTrimColor Then
    Set Screw = New clsFastener
    Screw.Quantity = Application.WorksheetFunction.RoundUp((((NetCornerLength / 30) * 2) + (NetRakeTrimLength / 12) + (NetRakeTrimLength / 30)) / 250, 0) * 250
    Screw.Color = b.OutsideCornerTrimColor
    TrimScrews.Add Screw
'add seperately if different colors
Else
    'rake trim screws
    Set Screw = New clsFastener
    Screw.Quantity = Application.WorksheetFunction.RoundUp(((NetRakeTrimLength / 12) + (NetRakeTrimLength / 30)) / 250, 0) * 250
    Screw.Color = b.RakeTrimColor
    TrimScrews.Add Screw
    'outside corner trim screws
    If NetCornerLength <> 0 Then
        Set Screw = New clsFastener
        Screw.Quantity = Application.WorksheetFunction.RoundUp(((NetCornerLength / 30) * 2) / 250, 0) * 250
        Screw.Color = b.OutsideCornerTrimColor
        TrimScrews.Add Screw
    End If
End If



End Sub


Private Sub SoffitScrewCalc(ScrewQty As Integer, SoffitScrewColor As String, SoffitType As String, b As clsBuilding)
Dim Location As String  'Wall Location
Dim pLines As Integer   'Purlin Lines


'determine wall location
Select Case True
Case InStr(1, SoffitType, "e1") <> 0
    Location = "e1"
Case InStr(1, SoffitType, "s2") <> 0
    Location = "s2"
Case InStr(1, SoffitType, "e3") <> 0
    Location = "e3"
Case InStr(1, SoffitType, "s4") <> 0
    Location = "s4"
End Select

'determine soffit screw color (only possibility since soffits of different color wouldn't be on the same building)
SoffitScrewColor = EstSht.Range(SoffitType).offset(0, 4).Value

With b
    'find roof purlin lines if needed
    If Location = "e1" Or Location = "e3" Then
        pLines = Application.WorksheetFunction.RoundUp((.RafterLength / 12) / 5, 0)
        'double purlins for gable roof
        If .rShape = "Gable" Then pLines = pLines * 2
        '''add in overhang/extensions
        'sidewall 2 roof purlins
        If .s2Extension = 0 Then
            'add in one purlin for an additional eave overhang
            If .s2Overhang > 4.25 Then pLines = pLines + 1
        ElseIf .s2Extension <> 0 Then
            pLines = pLines + Application.WorksheetFunction.RoundUp(((.s2ExtensionRafterLength + .s2ExtensionOverhangRafterLength) / 12) / 5, 0)
        End If
        'sidewall 4 roof purlins
        If .s4Extension = 0 Then
            'add in one purlin for an eave extension
            If .s4Overhang > 4.25 Then pLines = pLines + 1
        ElseIf .s4Extension <> 0 Then
            pLines = pLines + Application.WorksheetFunction.RoundUp(((.s4ExtensionRafterLength + .s4ExtensionOverhangRafterLength) / 12) / 5, 0)
        End If
    End If
        
        
    'determine extension/overhang type, calculate screws
    Select Case True
    Case InStr(1, SoffitType, "EaveOverhang") <> 0
        If (Location = "s2" And .s2ExtensionOverhang > 0) Or (Location = "s4" And .s4ExtensionOverhang > 0) Then
            'do nothing here I think
        Else
            'normal handling - just 2/ft along building length
            ScrewQty = ScrewQty + (.bLength * 2)
        End If
    Case InStr(1, SoffitType, "EaveExtension") <> 0
        'calculate extension purlin lines
        If Location = "s2" Then
            pLines = Application.WorksheetFunction.RoundUp(((.s2ExtensionRafterLength + .s2ExtensionOverhangRafterLength) / 12) / 5, 0)
        ElseIf Location = "s4" Then
            pLines = Application.WorksheetFunction.RoundUp(((.s4ExtensionRafterLength + .s4ExtensionOverhangRafterLength) / 12) / 5, 0)
        End If
        'screw/ft along purlin lines
        ScrewQty = ScrewQty + (.bLength * pLines)
    Case InStr(1, SoffitType, "GableOverhang") <> 0
        If Location = "e1" Then
            ScrewQty = ScrewQty + (((.e1Overhang + .e1ExtensionOverhang) / 12) * pLines)
        ElseIf Location = "e3" Then
            ScrewQty = ScrewQty + (((.e3Overhang + .e3ExtensionOverhang) / 12) * pLines)
        End If
    Case InStr(1, SoffitType, "GableExtension") <> 0
        If Location = "e1" Then
            ScrewQty = ScrewQty + ((.e1Extension / 12) * pLines)
        ElseIf Location = "e3" Then
            ScrewQty = ScrewQty + ((.e3Extension / 12) * pLines)
        End If
    End Select
End With


End Sub

Private Sub MatListSectionWrite(OutputSht As Worksheet, WriteCell As Range, MatCollection As Collection, CollectionType As String)
Dim Panel As clsPanel
Dim TrimPiece As clsTrim
Dim item As clsMiscItem
Dim StartCell As Range

'save start cell
Set StartCell = WriteCell

Select Case CollectionType
Case "Panel"
    For Each Panel In MatCollection
        If WriteCell <> StartCell Then OutputSht.Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = Panel.Quantity
        WriteCell.offset(0, 1).Value = Panel.PanelShape
        WriteCell.offset(0, 2).Value = Panel.PanelType
        WriteCell.offset(0, 3).Value = Panel.PanelMeasurement
        WriteCell.offset(0, 4).Value = Panel.PanelColor
        Set WriteCell = WriteCell.offset(1, 0)
    Next Panel
Case "Trim"
    For Each TrimPiece In MatCollection
        If WriteCell <> StartCell Then OutputSht.Rows(WriteCell.Row + 1).Insert
        WriteCell.Value = TrimPiece.Quantity
        WriteCell.offset(0, 1).Value = TrimPiece.tShape
        WriteCell.offset(0, 2).Value = TrimPiece.tType
        WriteCell.offset(0, 3).Value = TrimPiece.tMeasurement
        WriteCell.offset(0, 4).Value = TrimPiece.Color
        Set WriteCell = WriteCell.offset(1, 0)
    Next TrimPiece
End Select

End Sub

Private Sub MiscMaterialCalc(ButylTapeQty As Integer, InsideClosureQty As Integer, OutsideClosureQty As Integer, b As clsBuilding, rOverlaps As Integer)

With b
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Butyl Tape
    ButylTapeQty = Application.WorksheetFunction.RoundUp(.RoofFtLength / 3, 0) + 1
    If .rShape = "Single Slope" Then
        'one total rafter length for s2
        ButylTapeQty = ButylTapeQty * ((.s2RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength + .s2ExtensionOverhangRafterLength + .s4ExtensionOverhangRafterLength) / 12)
    ElseIf .rShape = "Gable" Then
        'rafter length along s2 and s4
        ButylTapeQty = ButylTapeQty * (((.s2RafterSheetLength + .s2ExtensionRafterLength + .s2ExtensionOverhangRafterLength) / 12) + ((.s4RafterSheetLength + .s4ExtensionRafterLength + .s4ExtensionOverhangRafterLength) / 12))
        'additional lengths of tape for the top
        ButylTapeQty = ButylTapeQty + (.bLength * 2)
    End If
    'add tape for overlaps
    ButylTapeQty = ButylTapeQty + (rOverlaps * .bLength)
    'increase by 5%, round up to nearest 44' roll
    ButylTapeQty = Application.WorksheetFunction.RoundUp((ButylTapeQty * 1.05) / 44, 0)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Inside Closures & Outside Closures
    If .rShape = "Single Slope" Then
        If .WallStatus("s2") = "Include" Then
            InsideClosureQty = Application.WorksheetFunction.RoundUp(.bLength / 3, 0)
            OutsideClosureQty = Application.WorksheetFunction.RoundUp(.bLength / 3, 0)
        End If
        If .WallStatus("e1") = "Include" Then OutsideClosureQty = OutsideClosureQty + Application.WorksheetFunction.RoundUp(.bWidth / 3, 0)
        If .WallStatus("e3") = "Include" Then OutsideClosureQty = OutsideClosureQty + Application.WorksheetFunction.RoundUp(.bWidth / 3, 0)
    ElseIf .rShape = "Gable" Then
        If .WallStatus("s2") = "Include" Then InsideClosureQty = Application.WorksheetFunction.RoundUp(.bLength / 3, 0)
        If .WallStatus("s4") = "Include" Then InsideClosureQty = InsideClosureQty + Application.WorksheetFunction.RoundUp(.bLength / 3, 0)
        If .WallStatus("e1") = "Include" Then OutsideClosureQty = OutsideClosureQty + Application.WorksheetFunction.RoundUp(.bWidth / 3, 0)
        If .WallStatus("e3") = "Include" Then OutsideClosureQty = OutsideClosureQty + Application.WorksheetFunction.RoundUp(.bWidth / 3, 0)
    End If
End With
    
    
    

End Sub


'''''''''''''''''''''' Sidewall Panel Generation
Sub SidewallPanelGen(SidewallPanels As Collection, sWall As String, b As clsBuilding, Optional FullHeightLinerPanels As Boolean)
Dim Panel As clsPanel
Dim sP1 As clsPanel
Dim sP2 As clsPanel
Dim sP3 As clsPanel
Dim WainscotPanel As clsPanel
Dim p1Length As Double
Dim p2Length As Double
Dim p3Length As Double
Dim SpecialBottomPurlin As Boolean
Dim WainscotFtLength As Double
Dim FO As clsFO
Dim FOCutoutp1 As clsPanel
Dim FOCollection As Collection




With b
    'Check for Wainscot
    If .Wainscot(sWall) <> "None" Then
        Set WainscotPanel = New clsPanel
        WainscotPanel.PanelLength = CDbl(Left(.Wainscot(sWall), 2))
        'only use wainscot ft length when not doing liner panels
        If FullHeightLinerPanels = False Then WainscotFtLength = WainscotPanel.PanelLength / 12
    End If
    '''''''''''''''''''''''''''' Generate Sidewall Panels
    Select Case .WallStatus(sWall)
    Case "Exclude"
        Exit Sub
    Case "Include", "Partial"
        If .WallStatus(sWall) = "Partial" Then SpecialBottomPurlin = True
    ''''' s4 of a Single Slope handling
        If .rShape = "Single Slope" And sWall = "s4" Then
            '''If high side eave is under 42 Feet
            If ((.HighSideEaveHeight / 12) - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) <= 42 Then
                Set sP1 = New clsPanel
                sP1.PanelLength = .HighSideEaveHeight - ((WainscotFtLength + .LengthAboveFinishedFloor(sWall)) * 12)
                'FullHeightLinerPanels is only TRUE when this sub is being called specifically to gen full height liner panels
                If FullHeightLinerPanels = True Then sP1.PanelLength = sP1.PanelLength - 8
            '''If high side eave is over 42 Feet and less than or equal to 84 feet
            ''' Since the highest purlin under 42' is at 37' 3.5", max height for 2 is 37'3.5"+42'
            ElseIf ((.HighSideEaveHeight / 12) - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) > 42 And ((.HighSideEaveHeight / 12) - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) <= (79 + (3.5 / 12)) Then
                ' Panel #1
                Set sP1 = New clsPanel
                sP1.PanelLength = ClosestWallPurlin((.HighSideEaveHeight - ((WainscotFtLength + .LengthAboveFinishedFloor(sWall)) * 12)) / 2, 0, SpecialBottomPurlin)
                'add overlap
                sP1.PanelLength = sP1.PanelLength + 1.5
                ' Panel #2
                Set sP2 = New clsPanel
                sP2.PanelLength = (.HighSideEaveHeight - ((WainscotFtLength + .LengthAboveFinishedFloor(sWall)) * 12)) - sP1.PanelLength
                sP2.PanelLength = sP2.PanelLength + 1.5
                If FullHeightLinerPanels = True Then sP2.PanelLength = sP2.PanelLength - 8
            ''''''''''''''' if high side eave is over 79' 3.5"
            ElseIf ((.HighSideEaveHeight / 12) - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) > (79 + (3.5 / 12)) Then
                ' Panel #1
                Set sP1 = New clsPanel
                sP1.PanelLength = ClosestWallPurlin((.HighSideEaveHeight - ((WainscotFtLength + .LengthAboveFinishedFloor(sWall)) * 12)) / 3, 0, SpecialBottomPurlin)
                'add overlap
                sP1.PanelLength = sP1.PanelLength + 1.5
                ' Panel #2
                Set sP2 = New clsPanel
                sP2.PanelLength = ClosestWallPurlin(sP1.PanelLength + ((.HighSideEaveHeight - ((WainscotFtLength + .LengthAboveFinishedFloor(sWall)) * 12)) / 3), 0, SpecialBottomPurlin)
                'two overlaps
                sP2.PanelLength = sP2.PanelLength + 3
                ' Panel #3
                Set sP3 = New clsPanel
                sP3.PanelLength = (.HighSideEaveHeight - ((WainscotFtLength + .LengthAboveFinishedFloor(sWall)) * 12)) - sP1.PanelLength - sP2.PanelLength
                'overlap
                sP3.PanelLength = sP3.PanelLength + 1.5
                If FullHeightLinerPanels = True Then sP3.PanelLength = sP3.PanelLength - 8
            End If
''''''''''''''''''''normal handling for everything but s4 on a single slope
        Else
            '''If building height is under 42 Feet
            If (.bHeight - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) <= 42 Then
                Set sP1 = New clsPanel
                sP1.PanelLength = (.bHeight - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) * 12
                If FullHeightLinerPanels = True Then sP1.PanelLength = sP1.PanelLength - 8
            '''If building height is over 42 Feet
            ElseIf (.bHeight - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) > 42 Then
                ' Panel #1
                Set sP1 = New clsPanel
                If ClosestWallPurlin(((.bHeight - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) / 2) * 12, 0, SpecialBottomPurlin) > (42 * 12) Then
                    'if over 42, then find next closest below
                    sP1.PanelLength = ClosestWallPurlin((((.bHeight - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) / 2) * 12), -1, SpecialBottomPurlin)
                Else
                    'find closest sidewall purlin if divided in half
                    sP1.PanelLength = ClosestWallPurlin(((.bHeight - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) / 2) * 12, 0, SpecialBottomPurlin)
                End If
                'add overlap
                sP1.PanelLength = sP1.PanelLength + 1.5
                'Panel #2
                Set sP2 = New clsPanel
                sP2.PanelLength = ((.bHeight - .LengthAboveFinishedFloor(sWall) - WainscotFtLength) * 12) - sP1.PanelLength
                'overlap
                sP2.PanelLength = sP2.PanelLength + 1.5
                If FullHeightLinerPanels = True Then sP2.PanelLength = sP2.PanelLength - 8
            End If
        End If
    End Select
End With

'''''''' Add Quantities
If Not sP1 Is Nothing Then sP1.Quantity = Application.WorksheetFunction.RoundUp(b.bLength / 3, 0)
If Not sP2 Is Nothing Then sP2.Quantity = sP1.Quantity
If Not sP3 Is Nothing Then sP3.Quantity = sP1.Quantity

''''Modify the sidewall panel collection (which is currently a solid wall) by removing panels that are covering a qualifying framed opening.
'' Qualifying framed openings are: Overhead Doors greater than or equal to 7' in width
If Not sP1 Is Nothing Then
    'set correct wall fo collection
    If sWall = "s2" Then Set FOCollection = b.s2FOs Else Set FOCollection = b.s4FOs
    For Each FO In FOCollection
        If (FO.FOType = "OHDoor" Or FO.FOType = "MiscFO") And FO.Width >= 7 * 12 And FO.bEdgeHeight = 0 Then
            Set FOCutoutp1 = New clsPanel
            FOCutoutp1.Quantity = Application.WorksheetFunction.RoundUp(FO.Width / (3 * 12), 0) - 2 'Calculate number of panels to cut short
            FOCutoutp1.PanelLength = sP1.PanelLength - (FO.Height - (WainscotFtLength * 12)) 'calculate panel length from top of FO to the next panel
            'subtract cutout panel from sidewall panel 1 quantity
            sP1.Quantity = sP1.Quantity - FOCutoutp1.Quantity
            'If there is more than 1 sidewall panel required for a well, check that only an overlap section isn't being added
            If Not sP2 Is Nothing Then
                If FOCutoutp1.PanelLength > 1.5 Then SidewallPanels.Add FOCutoutp1
            Else
                SidewallPanels.Add FOCutoutp1
            End If
        End If
    Next FO
    'add modified sidewall panel 1 to sidewall panel collection
    SidewallPanels.Add sP1
End If
If Not sP2 Is Nothing Then SidewallPanels.Add sP2
If Not sP3 Is Nothing Then SidewallPanels.Add sP3

'add parameters
For Each Panel In SidewallPanels
    Panel.PanelMeasurement = ImperialMeasurementFormat(Panel.PanelLength)
    Panel.PanelShape = b.wPanelShape
    Panel.PanelType = b.wPanelType
    Panel.PanelColor = b.wPanelColor
Next Panel
'wainscot
If Not WainscotPanel Is Nothing And FullHeightLinerPanels = False Then
    WainscotPanel.PanelMeasurement = ImperialMeasurementFormat(WainscotPanel.PanelLength)
    WainscotPanel.Quantity = sP1.Quantity   'only add quantity of full length sidewall panels
    WainscotPanel.PanelColor = EstSht.Range(sWall & "_Wainscot").offset(0, 2).Value
    WainscotPanel.PanelType = EstSht.Range(sWall & "_Wainscot").offset(0, 1).Value
    WainscotPanel.PanelShape = b.wPanelShape
    SidewallPanels.Add WainscotPanel
End If
  
End Sub

Private Sub LinerPanelGen(LinerPanels As Collection, b As clsBuilding, Location As String)
Dim LinerPanel As clsPanel
'''' Full Height Panels
If b.LinerPanels(Location) = "Full Height" Then
    'roof liner panels
    If Location = "Roof" Then
        Call RoofPanelGen(LinerPanels, b.RafterLength - 8, 0, b.bLength, b.rShape)
        If b.rShape = "Gable" Then
            For Each LinerPanel In LinerPanels
                LinerPanel.Quantity = LinerPanel.Quantity * 2
            Next LinerPanel
        End If
    'wall liner panels
    Else
        Select Case Location
        Case "e1", "e3"
            Call EndwallPanelGen(LinerPanels, Location, b, True)
        Case "s2", "s4"
            Call SidewallPanelGen(LinerPanels, Location, b, True)
        End Select
    End If
'''' 8' Liner Panels
ElseIf b.LinerPanels(Location) = "8'" Then
    Select Case Location
    Case "e1", "e3"
        Set LinerPanel = New clsPanel
        LinerPanel.Quantity = Application.WorksheetFunction.RoundUp(b.bWidth / 3, 0)
        LinerPanel.PanelLength = 8 * 12
        LinerPanels.Add LinerPanel
    Case "s2", "s4"
        Set LinerPanel = New clsPanel
        LinerPanel.Quantity = Application.WorksheetFunction.RoundUp(b.bLength / 3, 0)
        LinerPanel.PanelLength = 8 * 12
        LinerPanels.Add LinerPanel
    End Select
Else
    Exit Sub
End If

'add liner panel parameters
For Each LinerPanel In LinerPanels
    LinerPanel.PanelMeasurement = ImperialMeasurementFormat(LinerPanel.PanelLength)
    LinerPanel.PanelShape = EstSht.Range(Location & "_LinerPanels").offset(0, 1).Value
    LinerPanel.PanelType = EstSht.Range(Location & "_LinerPanels").offset(0, 2).Value
    LinerPanel.PanelColor = EstSht.Range(Location & "_LinerPanels").offset(0, 3).Value
Next LinerPanel

End Sub
