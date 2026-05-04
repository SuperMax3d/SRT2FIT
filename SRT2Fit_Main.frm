VERSION 5.00
Begin VB.Form SRT2Fit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SRT2Fit"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text_Weight 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      ToolTipText     =   "Weight in g"
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox Text_Scale 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      ToolTipText     =   "Default 1"
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox Text_Offset 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Default 0. Unit in seconds. 4 * 60 * 60 + 5 Seems to work with Avata360"
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton Cmd_Process 
      Caption         =   "Convert!"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text_Folder 
      Height          =   375
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   "Path to the folder containing the SRT files to convert."
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "Weight:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Scale:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Offset:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "SRT2Fit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' DJI SRT to FIT Converter
' Single VB6 Module (.bas) - Can also be used as a Form-based project
' by pasting into a Form and adjusting references.
'
' Converts DJI drone SRT subtitle files (with GPS/telemetry) into Garmin
' FIT activity files suitable for action cam overlay tools (e.g. Garmin VIRB,
' DashWare, GoPro Player, etc.)
'
' Outputs per record:
'   - Timestamp (UTC)
'   - Latitude / Longitude
'   - Altitude (absolute, meters)
'   - Speed (m/s, computed via Haversine between frames)
'   - Acceleration (m/s², derived from speed delta)
'   - Cumulative distance (meters)
'   - Estimated power (Watts) based on drone aerodynamic drag model
'   - Gimbal yaw / pitch / roll
'
' Time options:
'   - TimeOffsetSeconds  : shift all timestamps by N seconds (pos/neg)
'   - TimeScaleFactor    : scale elapsed time (e.g. 0.5 = slow motion source)
'
' FIT binary format written directly (no external SDK required).
' Complies with FIT Protocol v2.0 / Profile 21.x
'   File type   : Activity (4)
'   Messages    : file_id(0), file_creator(49), event(21),
'                 record(20), lap(19), session(18), activity(34)
'
' Usage (from a form or Sub Main):
'   Call ConvertDJISRTtoFIT("C:\path\input.srt", "C:\path\output.fit", _
'                           droneWeightKg:=0.895, _
'                           timeOffsetSec:=0.0, timeScale:=1.0)
'==============================================================================
Option Explicit

Private expr As String
Private pos As Long


' --- Data structures ---------------------------------------------------------

' Helper UDTs removed - LSet byte extraction is unreliable when VB6 pads
' Byte arrays inside UDTs. Use direct two's-complement byte extraction instead.
' For negative Longs, we decompose via unsigned arithmetic using a Double
' intermediate to avoid VB6's signed integer-division truncation bug.

Private Type DJIFrame
    FrameNum    As Long
    TimestampMs As Double   ' ms since video start (SRT time)
    UTCms       As Double   ' ms since Unix epoch (from embedded date string)
    Lat         As Double
    Lon         As Double
    RelAlt      As Double
    AbsAlt      As Double
    GbYaw       As Double
    GbPitch     As Double
    GbRoll      As Double
End Type

Private Type FITRecord
    TimestampUTC As Long    ' seconds since FIT epoch (1989-12-31 00:00:00 UTC)
    TimestampFrac As Double ' fractional second (0-1)
    Lat          As Long    ' semicircles (2^31 / 180 * deg)
    Lon          As Long    ' semicircles
    Altitude     As Long    ' (m + 500) * 5  encoded as per FIT spec
    Speed        As Long    ' mm/s
    Distance     As Long    ' cm
    Power        As Long    ' Watts (unsigned, use 0xFFFF for invalid)
    Accel        As Double  ' m/s² (stored as enhanced_speed trick; see notes)
    GbYaw        As Integer ' degrees * 10
    GbPitch      As Integer
    GbRoll       As Integer
    Valid        As Boolean
End Type



Public Function ListFiles(ByVal folderPath As String, Filter As String) As String()
    Dim Files() As String
    Dim fileName As String
    Dim count As Long
    
    ' Ensure path ends with backslash
    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Initialize
    count = -1
    fileName = Dir$(folderPath & Filter)
    
    ' Loop through files
    Do While fileName <> ""
        count = count + 1
        ReDim Preserve Files(count)
        Files(count) = folderPath & fileName
        fileName = Dir$
    Loop
    
    ' Handle no files found
    If count = -1 Then
        ReDim Files(0)
        Files(0) = ""
    End If
    
    ListFiles = Files
End Function

' --- Public entry point -----------------------------------------------------

Public Sub ConvertDJISRTtoFIT( _
        ByVal srtPath As String, _
        ByVal fitPath As String, _
        Optional ByVal droneWeightKg As Double = 0.895, _
        Optional ByVal timeOffsetSec As Double = 0, _
        Optional ByVal timeScaleFactor As Double = 1#)

    '-- 1. Parse SRT into frame array
    Dim frames() As DJIFrame
    Dim frameCount As Long
    ParseSRT srtPath, frames, frameCount
    If (frameCount = 0) Then
        MsgBox "No frames parsed from: " & srtPath, vbExclamation
        Exit Sub
    End If

    '-- 2. Compute derived channels (speed, accel, dist, power)
    Dim recs() As FITRecord
    BuildRecords frames, frameCount, droneWeightKg, _
                 timeOffsetSec, timeScaleFactor, recs

    '-- 3. Encode & write FIT binary
    Dim recCount As Long
    recCount = UBound(recs) + 1
    WriteFIT fitPath, recs, recCount

    'MsgBox "Done! FIT file written to:" & vbCrLf & fitPath, vbInformation
End Sub

' --- SRT Parser --------------------------------------------------------------

Private Sub ParseSRT(ByVal path As String, frames() As DJIFrame, cnt As Long)
    Dim Lines() As String
    Dim i As Long
    Dim j As Long
    Lines = Split(Replace$(ReadTxt(path), vbCrLf, vbLf), vbLf)

    ReDim frames(0 To 50000)
    cnt = 0

    Dim line1 As String, line2 As String, line3 As String, line4 As String
    Dim tagLine As String

    For i = 0 To UBound(Lines)
        ' Frame index line
        
        line1 = Trim(Lines(i))
        If Not IsNumeric(line1) Then GoTo SkipBlock

        ' Timecode line: 00:00:00,000 --> 00:00:00,016
        If (UBound(Lines) < i + 1) Then Exit For
        line2 = Trim(Lines(i + 1))

        ' First content line (FrameCnt / date)
        If (UBound(Lines) < i + 2) Then Exit For
        line3 = Trim(Lines(i + 2))
        line3 = StripFontTag(line3)

        ' Second content line (datetime)
        If (UBound(Lines) < i + 3) Then Exit For
        line4 = Trim(Lines(i + 3))    ' 2026-04-20 18:48:29.733

        ' Third content line (telemetry)
        tagLine = ""
        Dim extraLine As String
        j = 4
        Do While (UBound(Lines) >= i + j)
            extraLine = Trim(Lines(i + j))
            If extraLine = "" Then Exit Do
            tagLine = tagLine & extraLine
            j = j + 1
        Loop

        tagLine = StripFontTag(tagLine)

        ' Build frame
        Dim f As DJIFrame
        f.FrameNum = CLng(line1)

        ' SRT timestamp from timecode start
        Dim tcParts() As String
        tcParts = Split(line2, " --> ")
        f.TimestampMs = ParseTimecodeMs(tcParts(0))

        ' UTC from embedded date string (trim leading whitespace)
        f.UTCms = ParseDateTimeMs(Trim(line4))

        ' Parse telemetry tags
        f.Lat = ExtractTagDouble(tagLine, "latitude")
        f.Lon = ExtractTagDouble(tagLine, "longitude")
        f.RelAlt = ExtractTagDouble(tagLine, "rel_alt")
        f.AbsAlt = ExtractTagDouble(tagLine, "abs_alt")
        f.GbYaw = ExtractTagDouble(tagLine, "gb_yaw")
        f.GbPitch = ExtractTagDouble(tagLine, "gb_pitch")
        f.GbRoll = ExtractTagDouble(tagLine, "gb_roll")

        frames(cnt) = f
        cnt = cnt + 1
        If cnt > UBound(frames) - 1 Then
            ReDim Preserve frames(0 To UBound(frames) + 10000)
        End If

SkipBlock:
    Next
    ReDim Preserve frames(0 To cnt - 1)

End Sub
' --- Derived channel builder -------------------------------------------------
'
' WHY GPS INTERPOLATION IS NEEDED:
'   DJI video runs at 60 fps (16 ms/frame) but the GPS module updates at ~1 Hz.
'   This means ~60 consecutive frames share identical coordinates. Without
'   interpolation every frame-to-frame Haversine call returns 0, giving speed=0.
'
' STRATEGY:
'   Pass 1 – scan all frames and pre-compute the interpolated speed for every
'            frame by looking ahead to the next GPS change point, computing the
'            actual displacement over that GPS interval, then spreading that
'            speed evenly across all frames that fell inside the interval.
'   Pass 2 – downsample to ONE FIT record per unique whole-second timestamp.
'            FIT parsers require monotonically increasing integer timestamps.
'            Emitting 60 records with the same second value confuses every viewer.
'            We keep the last frame of each second (most up-to-date position).
'
' POWER MODEL:
'   P_hover = m*g * sqrt(m*g / (2*rho*A_disk))   induced hover power
'   P_drag  = 0.5 * rho * Cd * A * vH^3          aerodynamic drag
'   P_climb = m*g * vV                            climb work (positive only)

Private Sub BuildRecords( _
        frames() As DJIFrame, frameCount As Long, _
        droneKg As Double, offsetSec As Double, scaleT As Double, _
        recs() As FITRecord)

    Const RHO  As Double = 1.225
    Const CD   As Double = 0.4
    Const A    As Double = 0.04
    Const G    As Double = 9.80665
    Const PI   As Double = 3.14159265358979
    Const SC   As Double = 2147483648# / 180#   ' degrees -> semicircles
    Const FIT_EPOCH_OFFSET As Long = 631065600  ' Unix epoch -> FIT epoch (s)

    Const A_DISK As Double = 4# * 3.14159265 * 0.07 * 0.07

    Dim P_hover As Double
    P_hover = droneKg * G * Sqr(droneKg * G / (2# * RHO * A_DISK))

    ' -- Pass 1: compute adjusted timestamps and interpolated speed per frame --

    ' adjUTCms(i) = adjusted UTC timestamp in ms for each frame
    Dim adjUTCms() As Double
    ReDim adjUTCms(0 To frameCount - 1)

    Dim i As Long
    For i = 0 To frameCount - 1
        Dim elapsedRaw As Double
        elapsedRaw = frames(i).UTCms - frames(0).UTCms
        adjUTCms(i) = frames(0).UTCms + elapsedRaw * scaleT + offsetSec * 1000#
    Next i

    ' interpSpeedMs(i) = GPS-interval speed assigned to frame i (m/s)
    ' interpVV(i)      = vertical speed for frame i (m/s)
    ' cumDist(i)       = cumulative 3D distance at frame i (m)
    Dim interpSpeedMs() As Double
    Dim interpVV()      As Double
    Dim cumDistM()      As Double
    ReDim interpSpeedMs(0 To frameCount - 1)
    ReDim interpVV(0 To frameCount - 1)
    ReDim cumDistM(0 To frameCount - 1)

    Dim totalDist As Double: totalDist = 0
    Dim segStart  As Long:   segStart = 0    ' index of last GPS change

    ' Max plausible drone speed (m/s). DJI drones top out ~25 m/s.
    ' Any interpolated segment speed above this threshold is a GPS glitch
    ' (e.g. lat/lon briefly reads 0,0 = null island) and must be discarded.
    Const MAX_SPEED_MS As Double = 50#   ' generous ceiling: 180 km/h

    i = 1
    Do While i <= frameCount - 1
        Dim latChanged As Boolean
        latChanged = (frames(i).Lat <> frames(segStart).Lat) Or _
                     (frames(i).Lon <> frames(segStart).Lon) Or _
                     (frames(i).AbsAlt <> frames(segStart).AbsAlt)

        If latChanged Or (i = frameCount - 1) Then
            Dim segEnd As Long
            If latChanged Then
                segEnd = i - 1
            Else
                segEnd = i
            End If

            Dim nextIdx As Long
            If latChanged Then nextIdx = i Else nextIdx = i

            ' Skip any segment where either endpoint is 0,0 (GPS null island /
            ' signal loss). Haversine from real coords to 0,0 is ~9,600 km and
            ' produces impossible speeds that overflow CLng downstream.
            Dim startLat As Double: startLat = frames(segStart).Lat
            Dim startLon As Double: startLon = frames(segStart).Lon
            Dim endLat   As Double: endLat = frames(nextIdx).Lat
            Dim endLon   As Double: endLon = frames(nextIdx).Lon

            Dim validGPS As Boolean
            validGPS = (startLat <> 0# Or startLon <> 0#) And _
                       (endLat <> 0# Or endLon <> 0#)

            Dim dH As Double
            Dim dV As Double
            Dim d3D As Double
            Dim segSpeed3 As Double
            Dim segSpeedV As Double

            If validGPS Then
                dH = Haversine(startLat, startLon, endLat, endLon)
                dV = frames(nextIdx).AbsAlt - frames(segStart).AbsAlt
                d3D = Sqr(dH * dH + dV * dV)

                Dim segDtSec As Double
                segDtSec = (adjUTCms(nextIdx) - adjUTCms(segStart)) / 1000#
                If segDtSec <= 0# Then segDtSec = 0.001

                segSpeed3 = d3D / segDtSec
                segSpeedV = dV / segDtSec

                ' Sanity-clamp: if computed speed exceeds the physical maximum,
                ' this is a GPS glitch — treat the segment as stationary.
                If segSpeed3 > MAX_SPEED_MS Then
                    segSpeed3 = 0#: segSpeedV = 0#: d3D = 0#
                End If
            Else
                ' GPS unavailable for this segment — treat as stationary
                dH = 0#: dV = 0#: d3D = 0#: segSpeed3 = 0#: segSpeedV = 0#
            End If

            ' Distribute distance and speed evenly across every frame in segment
            Dim nFrames As Long: nFrames = nextIdx - segStart
            If nFrames < 1 Then nFrames = 1
            Dim distPerFrame As Double: distPerFrame = d3D / nFrames

            Dim j As Long
            For j = segStart To nextIdx - 1
                interpSpeedMs(j) = segSpeed3
                interpVV(j) = segSpeedV
                totalDist = totalDist + distPerFrame
                cumDistM(j) = totalDist
            Next j

            segStart = i
        End If
        i = i + 1
    Loop
    ' Last frame gets same as previous
    If frameCount > 0 Then
        interpSpeedMs(frameCount - 1) = interpSpeedMs(frameCount - 2)
        interpVV(frameCount - 1) = interpVV(frameCount - 2)
        cumDistM(frameCount - 1) = totalDist
    End If

    ' -- Pass 2: build FIT records, one per unique whole second --------------

    ReDim recs(0 To frameCount - 1)   ' over-allocate, trim later
    Dim recCount As Long: recCount = 0

    Dim prevFitTs   As Long:   prevFitTs = -1
    Dim prevSpeedMs As Double: prevSpeedMs = 0

    For i = 0 To frameCount - 1
        Dim unixSec As Double
        unixSec = adjUTCms(i) / 1000#
        Dim fitTs As Long
        fitTs = CLng(Int(unixSec)) - FIT_EPOCH_OFFSET

        ' Skip frames that share the same whole second as the previous emitted record,
        ' UNLESS this is the last frame of that second (peek ahead).
        ' This ensures exactly one FIT record per second with the freshest data.
        Dim nextFitTs As Long
        If i < frameCount - 1 Then
            Dim nextUnix As Double
            nextUnix = adjUTCms(i + 1) / 1000#
            nextFitTs = CLng(Int(nextUnix)) - FIT_EPOCH_OFFSET
        Else
            nextFitTs = fitTs + 1   ' force emit on last frame
        End If

        Dim doEmit As Boolean
        doEmit = True
        If fitTs = prevFitTs Then doEmit = False        ' same second as last emitted
        If nextFitTs = fitTs Then doEmit = False        ' not yet the last in this second

        If doEmit Then
            ' Smoothed speed (EMA)
            Dim rawSpeed As Double: rawSpeed = interpSpeedMs(i)
            Const ALPHA As Double = 0.4
            Dim smoothSpd As Double
            smoothSpd = ALPHA * rawSpeed + (1# - ALPHA) * prevSpeedMs

            Dim r As FITRecord

            r.TimestampUTC = fitTs

            ' Position in semicircles (SINT32)
            r.Lat = CLng(frames(i).Lat * SC)
            r.Lon = CLng(frames(i).Lon * SC)

            ' Altitude: FIT UINT16 = (m + 500) * 5
            r.Altitude = CLng((frames(i).AbsAlt + 500#) * 5#)

            ' Speed mm/s as UINT16 — clamp to UINT16 max (65535 = 235 km/h)
            Dim speedMms As Double: speedMms = smoothSpd * 1000#
            If speedMms > 65535# Then speedMms = 65535#
            If speedMms < 0# Then speedMms = 0#
            r.Speed = CLng(speedMms)

            ' Distance cm as UINT32 — clamp to Long max to avoid CLng overflow
            Dim distCm As Double: distCm = cumDistM(i) * 100#
            If distCm > 2147483647# Then distCm = 2147483647#
            If distCm < 0# Then distCm = 0#
            r.Distance = CLng(distCm)

            ' Acceleration m/s²
            r.Accel = smoothSpd - prevSpeedMs

            ' Power
            Dim vH As Double: vH = smoothSpd
            Dim vV As Double: vV = interpVV(i)
            Dim P_drag  As Double: P_drag = 0.5 * RHO * CD * A * vH * vH * vH
            Dim P_climb As Double
            If vV > 0# Then P_climb = droneKg * G * vV Else P_climb = 0#
            Dim P_total As Double: P_total = P_hover + P_drag + P_climb
            If P_total < 0# Then P_total = 0#
            If P_total > 32767# Then P_total = 32767#
            r.Power = CLng(P_total)

            r.GbYaw = CInt(frames(i).GbYaw * 10#)
            r.GbPitch = CInt(frames(i).GbPitch * 10#)
            r.GbRoll = CInt(frames(i).GbRoll * 10#)

            r.Valid = True
            recs(recCount) = r
            recCount = recCount + 1

            prevFitTs = fitTs
            prevSpeedMs = smoothSpd
        End If
    Next i

    ' Trim to actual record count
    If recCount > 0 Then
        ReDim Preserve recs(0 To recCount - 1)
    End If
End Sub

' --- FIT Binary Writer -------------------------------------------------------
'
' FIT file structure:
'   [File Header 14 bytes]
'   [Definition Message: file_id   (local 0)]
'   [Data    Message: file_id]
'   [Definition Message: event     (local 1)]
'   [Data    Message: event - start]
'   [Definition Message: record    (local 2)]
'   [Data    Message: record] * N
'   [Definition Message: lap       (local 3)]
'   [Data    Message: lap]
'   [Definition Message: session   (local 4)]
'   [Data    Message: session]
'   [Definition Message: activity  (local 5)]
'   [Data    Message: activity]
'   [CRC 2 bytes]
'
' All integers are Little-Endian unless noted.
Private Sub WriteFIT(ByVal fitPath As String, recs() As FITRecord, recCount As Long)

    Dim buf() As Byte
    ReDim buf(0 To 2000000)
    Dim pos As Long: pos = 14   ' reserve header

    Dim startTs As Long: startTs = recs(0).TimestampUTC
    Dim endTs   As Long: endTs = recs(recCount - 1).TimestampUTC
    Dim totalDistCm As Long: totalDistCm = recs(recCount - 1).Distance

    ' elapsed time in ms (FIT total_elapsed_time / total_timer_time unit = ms)
    Dim elapsedMs As Long
    elapsedMs = CLng((CDbl(endTs) - CDbl(startTs)) * 1000#)

    ' Field temp arrays (3 bytes: field_def_num, size, base_type)
    Dim f0(0 To 2) As Byte, f1(0 To 2) As Byte, f2(0 To 2) As Byte
    Dim f3(0 To 2) As Byte, f4(0 To 2) As Byte, f5(0 To 2) As Byte
    Dim f6(0 To 2) As Byte
    Dim flds(0 To 6) As Variant

    '   [Data    Messa    ' -- All definition messages use flat Byte arrays to avoid VB6 Variant
    ' aliasing: assigning lapX() into a Variant array stores a reference,
    ' not a copy. When g0..g11 are later assigned they overwrite the same
    ' module-level storage as lap0..lap9, corrupting the lap definition.
    ' Flat Byte arrays declared separately for each message have no aliasing.

    '----------------------------------------------------------------------
    ' file_id  (global 0, local 0)
    '----------------------------------------------------------------------
    Dim dFileId(0 To 11) As Byte
    dFileId(0) = 0: dFileId(1) = 1: dFileId(2) = 0 ' type ENUM
    dFileId(3) = 1: dFileId(4) = 2: dFileId(5) = 132 ' manufacturer UINT16
    dFileId(6) = 2: dFileId(7) = 2: dFileId(8) = 132 ' product UINT16
    dFileId(9) = 4: dFileId(10) = 4: dFileId(11) = 134 ' time_created UINT32
    WriteDefMsg buf, pos, 0, 0, dFileId, 4
    WriteDataHdr buf, pos, 0
    WriteByte buf, pos, 4                 ' type=activity(4)
    WriteUINT16 buf, pos, 255             ' manufacturer=development
    WriteUINT16 buf, pos, 0              ' product
    WriteUINT32 buf, pos, CLng(startTs)  ' time_created

    '----------------------------------------------------------------------
    ' event  (global 21, local 1) - timer start
    '----------------------------------------------------------------------
    Dim dEvent(0 To 8) As Byte
    dEvent(0) = 253: dEvent(1) = 4: dEvent(2) = 134 ' timestamp UINT32
    dEvent(3) = 0: dEvent(4) = 1: dEvent(5) = 0 ' event ENUM
    dEvent(6) = 1: dEvent(7) = 1: dEvent(8) = 0 ' event_type ENUM
    WriteDefMsg buf, pos, 1, 21, dEvent, 3
    WriteDataHdr buf, pos, 1
    WriteUINT32 buf, pos, CLng(startTs)
    WriteByte buf, pos, 0     ' event=timer(0)
    WriteByte buf, pos, 0     ' event_type=start(0)

    '----------------------------------------------------------------------
    ' record  (global 20, local 2)
    '----------------------------------------------------------------------
    Dim dRec(0 To 20) As Byte
    dRec(0) = 253: dRec(1) = 4: dRec(2) = 134 ' timestamp UINT32
    dRec(3) = 0: dRec(4) = 4: dRec(5) = 133 ' position_lat SINT32
    dRec(6) = 1: dRec(7) = 4: dRec(8) = 133 ' position_long SINT32
    dRec(9) = 2: dRec(10) = 2: dRec(11) = 132 ' altitude UINT16
    dRec(12) = 6: dRec(13) = 2: dRec(14) = 132 ' speed UINT16
    dRec(15) = 5: dRec(16) = 4: dRec(17) = 134 ' distance UINT32
    dRec(18) = 7: dRec(19) = 2: dRec(20) = 132 ' power UINT16
    WriteDefMsg buf, pos, 2, 20, dRec, 7

    Dim i As Long
    Dim r As FITRecord
    For i = 0 To recCount - 1
        r = recs(i)
        WriteDataHdr buf, pos, 2
        WriteUINT32 buf, pos, CLng(r.TimestampUTC)
        WriteSINT32 buf, pos, r.Lat
        WriteSINT32 buf, pos, r.Lon
        WriteUINT16 buf, pos, CInt(r.Altitude And &HFFFF&)
        WriteUINT16 buf, pos, CInt(r.Speed And &HFFFF&)
        WriteUINT32 buf, pos, r.Distance
        WriteUINT16 buf, pos, CInt(r.Power And &HFFFF&)
    Next i

    '----------------------------------------------------------------------
    ' event - timer stop (reuse local 1 definition)
    '----------------------------------------------------------------------
    WriteDataHdr buf, pos, 1
    WriteUINT32 buf, pos, CLng(endTs)
    WriteByte buf, pos, 0     ' event=timer(0)
    WriteByte buf, pos, 4     ' event_type=stop_disable(4)

    '----------------------------------------------------------------------
    ' lap  (global 19, local 3)
    '----------------------------------------------------------------------
    Dim dLap(0 To 29) As Byte
    dLap(0) = 254: dLap(1) = 2: dLap(2) = 132 ' message_index UINT16
    dLap(3) = 253: dLap(4) = 4: dLap(5) = 134 ' timestamp UINT32
    dLap(6) = 2: dLap(7) = 4: dLap(8) = 134 ' start_time UINT32
    dLap(9) = 3: dLap(10) = 4: dLap(11) = 133 ' start_position_lat SINT32
    dLap(12) = 4: dLap(13) = 4: dLap(14) = 133 ' start_position_long SINT32
    dLap(15) = 7: dLap(16) = 4: dLap(17) = 134 ' total_elapsed_time UINT32
    dLap(18) = 8: dLap(19) = 4: dLap(20) = 134 ' total_timer_time UINT32
    dLap(21) = 9: dLap(22) = 4: dLap(23) = 134 ' total_distance UINT32
    dLap(24) = 0: dLap(25) = 1: dLap(26) = 0 ' event ENUM
    dLap(27) = 1: dLap(28) = 1: dLap(29) = 0 ' event_type ENUM
    WriteDefMsg buf, pos, 3, 19, dLap, 10
    WriteDataHdr buf, pos, 3
    WriteUINT16 buf, pos, 0
    WriteUINT32 buf, pos, CLng(endTs)
    WriteUINT32 buf, pos, CLng(startTs)
    WriteSINT32 buf, pos, recs(0).Lat
    WriteSINT32 buf, pos, recs(0).Lon
    WriteUINT32 buf, pos, elapsedMs
    WriteUINT32 buf, pos, elapsedMs
    WriteUINT32 buf, pos, totalDistCm
    WriteByte buf, pos, 9     ' event=lap(9)
    WriteByte buf, pos, 1     ' event_type=stop(1)

    '----------------------------------------------------------------------
    ' session  (global 18, local 4)
    '----------------------------------------------------------------------
    Dim dSess(0 To 35) As Byte
    dSess(0) = 254: dSess(1) = 2: dSess(2) = 132 ' message_index UINT16
    dSess(3) = 253: dSess(4) = 4: dSess(5) = 134 ' timestamp UINT32
    dSess(6) = 2:  dSess(7) = 4: dSess(8) = 134 ' start_time UINT32
    dSess(9) = 3:  dSess(10) = 4: dSess(11) = 133 ' start_position_lat SINT32
    dSess(12) = 4: dSess(13) = 4: dSess(14) = 133 ' start_position_long SINT32
    dSess(15) = 7: dSess(16) = 4: dSess(17) = 134 ' total_elapsed_time UINT32
    dSess(18) = 8: dSess(19) = 4: dSess(20) = 134 ' total_timer_time UINT32
    dSess(21) = 9: dSess(22) = 4: dSess(23) = 134 ' total_distance UINT32
    dSess(24) = 0: dSess(25) = 1: dSess(26) = 0 ' event ENUM
    dSess(27) = 1: dSess(28) = 1: dSess(29) = 0 ' event_type ENUM
    dSess(30) = 25: dSess(31) = 2: dSess(32) = 132 ' first_lap_index UINT16
    dSess(33) = 26: dSess(34) = 2: dSess(35) = 132 ' num_laps UINT16
    WriteDefMsg buf, pos, 4, 18, dSess, 12
    WriteDataHdr buf, pos, 4
    WriteUINT16 buf, pos, 0
    WriteUINT32 buf, pos, CLng(endTs)
    WriteUINT32 buf, pos, CLng(startTs)
    WriteSINT32 buf, pos, recs(0).Lat
    WriteSINT32 buf, pos, recs(0).Lon
    WriteUINT32 buf, pos, elapsedMs
    WriteUINT32 buf, pos, elapsedMs
    WriteUINT32 buf, pos, totalDistCm
    WriteByte buf, pos, 9     ' event=session(9)
    WriteByte buf, pos, 1     ' event_type=stop(1)
    WriteUINT16 buf, pos, 0   ' first_lap_index
    WriteUINT16 buf, pos, 1   ' num_laps

    '----------------------------------------------------------------------
    ' activity  (global 34, local 5)
    '----------------------------------------------------------------------
    Dim dAct(0 To 20) As Byte
    dAct(0) = 253: dAct(1) = 4: dAct(2) = 134 ' timestamp UINT32
    dAct(3) = 0: dAct(4) = 4: dAct(5) = 134 ' total_timer_time UINT32
    dAct(6) = 1: dAct(7) = 2: dAct(8) = 132 ' num_sessions UINT16
    dAct(9) = 2: dAct(10) = 1: dAct(11) = 0 ' type ENUM
    dAct(12) = 3: dAct(13) = 1: dAct(14) = 0 ' event ENUM
    dAct(15) = 4: dAct(16) = 1: dAct(17) = 0 ' event_type ENUM
    dAct(18) = 5: dAct(19) = 4: dAct(20) = 134 ' local_timestamp UINT32
    WriteDefMsg buf, pos, 5, 34, dAct, 7
    WriteDataHdr buf, pos, 5
    WriteUINT32 buf, pos, CLng(endTs)
    WriteUINT32 buf, pos, elapsedMs
    WriteUINT16 buf, pos, 1
    WriteByte buf, pos, 0      ' type=manual(0)
    WriteByte buf, pos, 26     ' event=activity(26)
    WriteByte buf, pos, 1      ' event_type=stop(1)
    WriteUINT32 buf, pos, CLng(endTs)   ' local_timestamp


    '----------------------------------------------------------------------
    ' File Header (14 bytes): fill in now we know total data size
    '----------------------------------------------------------------------
    Dim dataSize As Long
    dataSize = pos - 14   ' excludes header and trailing CRC

    buf(0) = 14            ' header size
    buf(1) = 16            ' protocol version 1.0 encoded as &H10
    buf(2) = 8             ' profile version LSB  (21.08 -> &H0854 LE)
    buf(3) = 84            ' profile version MSB
    ' data size UINT32 LE — dataSize is always positive, safe to use Double Mod
    Dim dsD As Double: dsD = CDbl(dataSize)
    buf(4) = CByte(CLng(dsD Mod 256#))
    buf(5) = CByte(CLng(Int(dsD / 256#) Mod 256#))
    buf(6) = CByte(CLng(Int(dsD / 65536#) Mod 256#))
    buf(7) = CByte(CLng(Int(dsD / 16777216#) Mod 256#))
    ' .FIT magic
    buf(8) = Asc(".")
    buf(9) = Asc("F")
    buf(10) = Asc("I")
    buf(11) = Asc("T")
    ' Header CRC (CRC16 of bytes 0-11)
    Dim hdrCRC As Long
    hdrCRC = CalcCRC16(buf, 0, 11)
    buf(12) = CByte(CLng(CDbl(hdrCRC) Mod 256#))
    buf(13) = CByte(CLng(Int(CDbl(hdrCRC) / 256#) Mod 256#))

    '----------------------------------------------------------------------
    ' File CRC (last 2 bytes, CRC16 of entire content so far)
    '----------------------------------------------------------------------
    Dim fileCRC As Long
    fileCRC = CalcCRC16(buf, 0, pos - 1)
    buf(pos) = CByte(CLng(CDbl(fileCRC) Mod 256#)):                      pos = pos + 1
    buf(pos) = CByte(CLng(Int(CDbl(fileCRC) / 256#) Mod 256#)):         pos = pos + 1

    '----------------------------------------------------------------------
    ' Write to disk — Kill first so VB6 binary Open truncates the file
    '----------------------------------------------------------------------
    Dim fNum As Integer
    fNum = FreeFile
    On Error Resume Next
    Kill fitPath          ' delete any previous version; ignore error if not found
    On Error GoTo 0
    Open fitPath For Binary Access Write As #fNum
    Dim finalBuf() As Byte
    ReDim finalBuf(0 To pos - 1)
    Dim b As Long
    For b = 0 To pos - 1
        finalBuf(b) = buf(b)
    Next b
    Put #fNum, 1, finalBuf
    Close #fNum
End Sub

' --- FIT buffer helpers -------------------------------------------------------


' Write a FIT Definition Message directly into buf.
' fieldDefs is a flat array of (field_def_num, size, base_type) triplets,
' laid out as fieldDefs(0)=f0_num, fieldDefs(1)=f0_size, fieldDefs(2)=f0_type,
'            fieldDefs(3)=f1_num, fieldDefs(4)=f1_size, fieldDefs(5)=f1_type, ...
' nFields = number of fields (array must have nFields*3 elements, 0-based).
' This avoids the Variant array aliasing bug where VB6 stores references
' (not copies) into Variant arrays, causing later assignments to corrupt
' earlier entries when the underlying Byte arrays share module-level storage.
Private Sub WriteDefMsg(buf() As Byte, pos As Long, _
        localNum As Integer, globalNum As Integer, _
        fieldDefs() As Byte, nFields As Integer)

    buf(pos) = CByte(&H40 Or (localNum And &HF)): pos = pos + 1
    buf(pos) = 0: pos = pos + 1       ' reserved
    buf(pos) = 0: pos = pos + 1       ' architecture: little-endian
    buf(pos) = CByte(globalNum And &HFF): pos = pos + 1
    buf(pos) = CByte((globalNum \ 256) And &HFF): pos = pos + 1
    buf(pos) = CByte(nFields): pos = pos + 1

    Dim i As Integer
    For i = 0 To nFields - 1
        buf(pos) = fieldDefs(i * 3):     pos = pos + 1  ' field def num
        buf(pos) = fieldDefs(i * 3 + 1): pos = pos + 1  ' size
        buf(pos) = fieldDefs(i * 3 + 2): pos = pos + 1  ' base type
    Next i
End Sub


Private Sub WriteDataHdr(buf() As Byte, pos As Long, localNum As Integer)
    buf(pos) = CByte(localNum And &HF): pos = pos + 1
End Sub

Private Sub WriteByte(buf() As Byte, pos As Long, val As Integer)
    buf(pos) = CByte(val And &HFF): pos = pos + 1
End Sub

Private Sub WriteUINT16(buf() As Byte, pos As Long, val As Integer)
    ' LSet reinterpretation avoids arithmetic on potentially negative Integer values.
    Dim iv As Integer
    Dim ib(0 To 1) As Byte
    ' VB6 has no Integer-to-2Bytes UDT trick as cleanly, so use safe masking:
    ' For UINT16 the value is always 0-65535 after the And &HFFFF& in callers,
    ' but val is Integer (signed). Use Long arithmetic to stay safe.
    Dim uval As Long
    uval = CLng(val) And &HFFFF&
    buf(pos) = CByte(uval And &HFF&)
    buf(pos + 1) = CByte((uval \ 256&) And &HFF&)
    pos = pos + 2
End Sub

Private Sub WriteUINT32(buf() As Byte, pos As Long, val As Long)
    ' Extract 4 LE bytes from a Long (signed) without any overflow.
    ' Strategy: split into two unsigned 16-bit halves using Double arithmetic,
    ' then extract each byte from the halves (each half fits safely in a Long).
    '
    ' For positive val:  hi16 = val \ 65536,  lo16 = val - hi16*65536
    ' For negative val:  use Double to get the unsigned 32-bit representation
    '   uv = val + 2^32  (always 0..4294967295 as Double)
    '   then split into hi16 = Int(uv / 65536), lo16 = uv - hi16*65536
    ' Both hi16 and lo16 are 0..65535 — they fit in Long with no overflow.
    Dim uv  As Double
    Dim hi  As Long    ' upper 16 bits (0..65535)
    Dim lo  As Long    ' lower 16 bits (0..65535)

    If val >= 0 Then
        uv = CDbl(val)
    Else
        uv = CDbl(val) + 4294967296#
    End If

    hi = CLng(Int(uv / 65536#))        ' 0..65535 — safe for Long
    lo = CLng(uv - CDbl(hi) * 65536#)  ' 0..65535 — safe for Long

    buf(pos) = CByte(lo And &HFF)
    buf(pos + 1) = CByte(lo \ 256)
    buf(pos + 2) = CByte(hi And &HFF)
    buf(pos + 3) = CByte(hi \ 256)
    pos = pos + 4
End Sub

Private Sub WriteSINT32(buf() As Byte, pos As Long, val As Long)
    ' Same binary layout as UINT32 for two's complement
    WriteUINT32 buf, pos, val
End Sub

' --- CRC-16 (FIT protocol lookup-table algorithm, per Garmin FIT spec) ----------
' The FIT file format uses a specific 16-entry nibble lookup table CRC,
' NOT standard CRC-CCITT. Table and algorithm taken directly from the
' Garmin FIT SDK documentation.

Private Function CalcCRC16(buf() As Byte, startPos As Long, endPos As Long) As Long
    ' Garmin FIT CRC table (16 entries, one per nibble)
    Dim tbl(0 To 15) As Long
    tbl(0) = &H0: tbl(1) = &HCC01: tbl(2) = &HD801: tbl(3) = &H1400
    tbl(4) = &HF001: tbl(5) = &H3C00: tbl(6) = &H2800: tbl(7) = &HE401
    tbl(8) = &HA001: tbl(9) = &H6C00: tbl(10) = &H7800: tbl(11) = &HB401
    tbl(12) = &H5000: tbl(13) = &H9C01: tbl(14) = &H8801: tbl(15) = &H4400

    Dim crc As Long: crc = 0
    Dim i As Long
    Dim b As Long
    Dim tmp As Long

    For i = startPos To endPos
        b = buf(i) And &HFF&

        ' Process low nibble
        tmp = tbl(crc And &HF&)
        crc = (crc And &HFFFF&) \ 16    ' logical right-shift by 4 (no sign ext)
        crc = crc Xor tbl(b And &HF&) Xor tmp

        ' Process high nibble
        tmp = tbl(crc And &HF&)
        crc = (crc And &HFFFF&) \ 16    ' logical right-shift by 4
        crc = crc Xor tbl((b And &HFF&) \ 16) Xor tmp
    Next i

    CalcCRC16 = crc And &HFFFF&
End Function


' --- Haversine distance (lat1,lon1 -> lat2,lon2) in meters --------------------

Private Function Haversine(lat1 As Double, lon1 As Double, _
                            lat2 As Double, lon2 As Double) As Double
    Const r As Double = 6371000#  ' Earth radius meters
    Const PI As Double = 3.14159265358979
    Dim dlat As Double: dlat = (lat2 - lat1) * PI / 180#
    Dim dlon As Double: dlon = (lon2 - lon1) * PI / 180#
    Dim A As Double
    A = Sin(dlat / 2) * Sin(dlat / 2) + _
        Cos(lat1 * PI / 180#) * Cos(lat2 * PI / 180#) * _
        Sin(dlon / 2) * Sin(dlon / 2)
    Haversine = r * 2 * Atn(Sqr(A) / Sqr(1 - A))
End Function

' --- SRT parsing helpers -----------------------------------------------------

' Convert timecode "HH:MM:SS,mmm" to milliseconds
Private Function ParseTimecodeMs(tc As String) As Double
    tc = Trim(tc)
    Dim parts() As String: parts = Split(tc, ":")
    Dim ss() As String: ss = Split(parts(2), ",")
    ParseTimecodeMs = CDbl(parts(0)) * 3600000 + _
                      CDbl(parts(1)) * 60000 + _
                      CDbl(ss(0)) * 1000 + _
                      CDbl(ss(1))
End Function

' Convert DJI embedded date "2026-04-20 18:48:29.733" to ms since Unix epoch
Private Function ParseDateTimeMs(dt As String) As Double
    dt = Trim(dt)
    If Len(dt) < 19 Then ParseDateTimeMs = 0: Exit Function

    Dim yr As Integer, mo As Integer, dy As Integer
    Dim hh As Integer, mm As Integer
    Dim ss As Double

    yr = CInt(Mid(dt, 1, 4))
    mo = CInt(Mid(dt, 6, 2))
    dy = CInt(Mid(dt, 9, 2))
    hh = CInt(Mid(dt, 12, 2))
    mm = CInt(Mid(dt, 15, 2))
    ss = CDbl(Mid(dt, 18))   ' e.g. "29.733"

    ' Julian Day Number for Gregorian calendar
    Dim A As Long, Y As Long, m As Long, jdn As Long
    A = (14 - mo) \ 12
    Y = yr + 4800 - A
    m = mo + 12 * A - 3
    jdn = dy + (153 * m + 2) \ 5 + 365 * Y + Y \ 4 - Y \ 100 + Y \ 400 - 32045

    ' Unix epoch JDN = 2440588 (1970-01-01)
    Dim daysSinceEpoch As Long
    daysSinceEpoch = jdn - 2440588

    ParseDateTimeMs = CDbl(daysSinceEpoch) * 86400000# + _
                      CDbl(hh) * 3600000# + _
                      CDbl(mm) * 60000# + _
                      ss * 1000#
End Function

' Extract numeric value from DJI tag string e.g. "[latitude: 45.386868]"
Private Function ExtractTagDouble(s As String, tag As String) As Double
    Dim p As Long
    p = InStr(1, s, "[" & tag & ":")
    If p = 0 Then ExtractTagDouble = 0: Exit Function
    p = p + Len("[" & tag & ":")
    ' Skip space
    Do While Mid(s, p, 1) = " ": p = p + 1: Loop
    ' Read until ] or space
    Dim numStr As String: numStr = ""
    Do While p <= Len(s)
        Dim c As String: c = Mid(s, p, 1)
        If c = "]" Or c = " " Then Exit Do
        numStr = numStr & c
        p = p + 1
    Loop
    On Error Resume Next
    ExtractTagDouble = CDbl(numStr)
    On Error GoTo 0
End Function

' Strip <font ...> and </font> tags from a line
Private Function StripFontTag(s As String) As String
    Dim result As String: result = s
    ' Remove opening font tag
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, LCase(result), "<font")
    If p1 > 0 Then
        p2 = InStr(p1, result, ">")
        If p2 > 0 Then result = Left(result, p1 - 1) & Mid(result, p2 + 1)
    End If
    ' Remove closing font tag
    p1 = InStr(1, LCase(result), "</font>")
    If p1 > 0 Then result = Left(result, p1 - 1) & Mid(result, p1 + 7)
    StripFontTag = result
End Function

' --- Optional: Sub Main entry point for standalone .exe ----------------------
Private Function SafeUbound(Arr() As String) As Long
    On Error Resume Next
    SafeUbound = -1
    SafeUbound = UBound(Arr)
End Function

Private Sub Form_Load()
    Dim i As Long
    Dim DataArr() As String
    
    DataArr = Split(ReadTxt(App.path & "\SRT2FIT.txt"), vbCrLf)
    If (SafeUbound(DataArr) > 2) Then
        Text_Folder.Text = DataArr(0)
        Text_Offset.Text = DataArr(1)
        Text_Scale.Text = DataArr(2)
        Text_Weight.Text = DataArr(3)
    Else
        Text_Folder.Text = App.path
        Text_Offset.Text = "0"
        Text_Scale.Text = "1"
        Text_Weight.Text = "0.249"
    End If
End Sub

Private Sub Cmd_Process_Click()
    On Error GoTo ErrorHandler
    Dim srtFile As String
    Dim fitFile As String
    Dim SkippedFiles As String
    Dim ProcessedCount As Long

    Dim weightKg As Double
    Dim offsetSec As Double
    Dim scaleF As Double
    Dim Files() As String
    Dim i As Long
    
    offsetSec = EvalExpr(Text_Offset.Text)
    scaleF = EvalExpr(Text_Scale.Text)
    weightKg = EvalExpr(Text_Weight.Text)
    
    ProcessedCount = 0
    Files = ListFiles(Text_Folder.Text, "*.SRT")
    
    If (SafeUbound(Files) > -1) Then
        
        For i = LBound(Files) To UBound(Files)
            If (Files(i) <> "") Then
                srtFile = Files(i)
                fitFile = Left(srtFile, InStrRev(srtFile, ".")) & "fit"
                ConvertDJISRTtoFIT srtFile, fitFile, weightKg, offsetSec, scaleF
                ProcessedCount = ProcessedCount + 1
            End If
        Next i
        
        If (SkippedFiles <> "") Then
            MsgBox "Conversion Completed!" & vbCrLf & "Processed Files: " & ProcessedCount & "Skipped the following files: " & SkippedFiles
        Else
            MsgBox "Conversion Completed!" & vbCrLf & "Processed Files: " & ProcessedCount
        End If
    Else
        MsgBox "No SRT file found in: " & Text_Folder.Text, vbCritical
    End If
    
    Exit Sub
ErrorHandler:
    SkippedFiles = SkippedFiles & vbCrLf & Files(i) & " (Error: " & Err.Desc & ")"
    ProcessedCount = ProcessedCount - 1
    Resume Next
End Sub

'--------------------------------
Public Sub WriteTxt(ByVal filePath As String, ByVal content As String)
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    
    If Len(content) > 0 Then
        Put #fileNum, , content
    End If
    
    Close #fileNum
End Sub

Public Function ReadTxt(ByVal filePath As String) As String
    Dim fileNum As Integer
    Dim fileLen As Long
    Dim fileData As String

    fileNum = FreeFile
    Open filePath For Binary As #fileNum

    fileLen = LOF(fileNum)
    If fileLen > 0 Then
        fileData = String$(fileLen, vbNullChar)
        Get #fileNum, , fileData
    End If

    Close #fileNum

    ReadTxt = fileData
End Function





'-------------------------------------
Public Function EvalExpr(ByVal s As String) As Double
    expr = Replace$(s, " ", "") ' remove spaces
    pos = 1
    EvalExpr = ParseAddSub
End Function

Private Function ParseAddSub() As Double
    Dim value As Double
    value = ParseMulDiv
    
    Do While pos <= Len(expr)
        Dim op As String
        op = Mid$(expr, pos, 1)
        
        If op <> "+" And op <> "-" Then Exit Do
        pos = pos + 1
        
        If op = "+" Then
            value = value + ParseMulDiv
        Else
            value = value - ParseMulDiv
        End If
    Loop
    
    ParseAddSub = value
End Function

Private Function ParseMulDiv() As Double
    Dim value As Double
    value = ParseNumber
    
    Do While pos <= Len(expr)
        Dim op As String
        op = Mid$(expr, pos, 1)
        
        If op <> "*" And op <> "/" Then Exit Do
        pos = pos + 1
        
        If op = "*" Then
            value = value * ParseNumber
        Else
            value = value / ParseNumber
        End If
    Loop
    
    ParseMulDiv = value
End Function

Private Function ParseNumber() As Double
    Dim startPos As Long
    startPos = pos
    
    Do While pos <= Len(expr) And _
        (Mid$(expr, pos, 1) Like "[0-9.]")
        pos = pos + 1
    Loop
    
    ParseNumber = CDbl(Mid$(expr, startPos, pos - startPos))
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim Str As String
    Str = Text_Folder.Text
    Str = Str & vbCrLf & Text_Offset.Text
    Str = Str & vbCrLf & Text_Scale.Text
    Str = Str & vbCrLf & Text_Weight.Text
    WriteTxt App.path & "\SRT2FIT.txt", Str
End Sub



Private Sub Text_Folder_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Check for Ctrl + A
    If (Shift And vbCtrlMask) <> 0 And KeyCode = vbKeyA Then
        Text_Folder.SelStart = 0
        Text_Folder.SelLength = Len(Text_Folder.Text)
        KeyCode = 0 ' prevent default behavior
    End If
End Sub

Private Sub Text_Folder_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim filePath As String
    
    On Error Resume Next
    
    ' Check if the data contains files/folders
    If Data.GetFormat(vbCFFiles) Then
        ' Take the first dropped item
        filePath = Data.Files(1)
        
        ' If it's a file, extract its folder
        If Dir$(filePath, vbDirectory) = "" Then
            ' It's a file ? strip filename
            filePath = Left$(filePath, InStrRev(filePath, "\") - 1)
        End If
        
        ' Set textbox
        Text_Folder.Text = filePath
    End If
End Sub


Private Sub Text_Offset_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Check for Ctrl + A
    If (Shift And vbCtrlMask) <> 0 And KeyCode = vbKeyA Then
        Text_Offset.SelStart = 0
        Text_Offset.SelLength = Len(Text_Folder.Text)
        KeyCode = 0 ' prevent default behavior
    End If
End Sub

Private Sub Text_Scale_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Check for Ctrl + A
    If (Shift And vbCtrlMask) <> 0 And KeyCode = vbKeyA Then
        Text_Scale.SelStart = 0
        Text_Scale.SelLength = Len(Text_Folder.Text)
        KeyCode = 0 ' prevent default behavior
    End If
End Sub

Private Sub Text_Weight_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Check for Ctrl + A
    If (Shift And vbCtrlMask) <> 0 And KeyCode = vbKeyA Then
        Text_Weight.SelStart = 0
        Text_Weight.SelLength = Len(Text_Folder.Text)
        KeyCode = 0 ' prevent default behavior
    End If
End Sub
