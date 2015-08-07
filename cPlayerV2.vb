Imports DirectShowLib
Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D
Imports System.Runtime.InteropServices

Namespace ESTRUCTURAS
    Public Structure UNDER_THE_HOOD
        Public FilterName As String
        Public Pins() As PIN_INFO
    End Structure

    Public Structure PIN_INFO
        Public PinName As String
        Public PinDirection As PinDirection
        Public MajorType As Guid
        Public SubType As Guid
        Public SampleSize As Long
    End Structure

    Public Structure VIDEO_STATS
        Public AVG_FRAME_RATE As Double
        Public AVG_SYNC_OFFSET As Long
        Public FRAMES_DRAWN As Long
        Public FRAMES_DROPPED As Long
    End Structure

    Public Structure AUDIO_STATS
        Public CHANNELS As Long
        Public SAMPLES_PER_SEC As Long
        Public AVG_BYTES_PER_SEC As Long
        Public BUFFER_DURATION As Long
        Public BUFFER_FILL As Long
    End Structure

    Public Structure M_M_V_D
        Public dwMinValue As Single
        Public dwMaxValue As Single
        Public dwStep As Single
        Public dwDefault As Single
        Public dwCurrrentValue As Single
        Public Sub New(ByVal Val As Single)
            Me.dwCurrrentValue = Val
        End Sub
    End Structure

    Public Structure TextOverlayParams
        Public FontFace As String
        Public FontColor As Color
        Public Bold As Boolean
        Public Italic As Boolean
        Public UnderLine As Boolean
        Public DrawBorder As Boolean
        Public AlphaColor As Color
        Public OutlineSize As Single
        Public FontSize As Single
    End Structure
End Namespace

Namespace ENUMS
    Public Enum ESTADO
        ESTADO_CERRADO
        ESTADO_ABIERTO
        ESTADO_REPRODUCIENDO
        ESTADO_PAUSADO
        ESTADO_DETENIDO
    End Enum

    Public Enum ERROR_LIST
        ERL_DIRECTSHOW
        ERL_D3DDEVICE
        ERL_D3DSURFACE
        ERL_RUTAINCORRECTA
        ERL_MODOINCORRECTO
        ERL_VALORINCORRECTO
        ERL_TFNOSOPORTADO
        ERL_VENTANAINEXISTENTE
        ERL_DESCONOCIDO
    End Enum

    Public Enum FILTER_TYPE
        FT_SOURCE
        FT_TRANSFORM
        FT_RENDERER
    End Enum

    Public Enum TIME_FORMAT
        TF_MILLISECONDS
        TF_FRAMES
        TF_SECONDS
        TF_MEDIATIME
        TF_SAMPLES
    End Enum

    Public Enum PROP
        BRIGHNESS
        HUE
        SATURATION
        CONTRAST
    End Enum

End Namespace

Namespace CLASES

    Public Class WIN32_API
        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True, ExactSpelling:=True)> _
        Public Shared Function IsWindow(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("user32.dll")> _
        Public Shared Function GetDC(ByVal hwnd As IntPtr) As IntPtr
        End Function
    End Class

    Public Class cPlayerV2
        Implements IDisposable

        'UTILIZADO PARA OBTENER UN PUNTERO NO ADMINISTRADO DEL SURFACE DE DIRECT3D
        Private Const DxMagicNumber As Integer = -759872593

#Region "EVENTOS"
        'NO UTILIZAREMOS EVENTOS CON SUSCRIPCION (HANDLERS) PORQUE ES UN BARDO LIBERARLOS Y NO LOS NECESITAMOS
        Public Event FileOpened()
        Public Event CustomGraphRendered()
        Public Event UsingIC()
        Public Event VolumeChanged()
        Public Event BalanceChanged()
        Public Event TimeFormatChanged()
        Public Event [Error](ByVal errType As ENUMS.ERROR_LIST)

#End Region

#Region "MIEMBROS DE INSTANCIA"
        'CAMPOS DE INSTANCIA
        Private mMainHWND As IntPtr 'USADO POR DIRECT3D
        Private mHWND As IntPtr
        Private StreamState As ENUMS.ESTADO
        Private mFileName As String
        Private mIsUsingIC As Boolean
        Private mHasAudio As Boolean
        Private mHasVideo As Boolean
        Private mTimeFormat As ENUMS.TIME_FORMAT
        Private mTextParams As ESTRUCTURAS.TextOverlayParams
        Private mVideoSize As Size
        Private mHasText As Boolean
#End Region

#Region "MIEMBROS NECESARIOS PARA HACER FUNCIONAR A DIRECT3D"
        'OBJETOS DE DIRECT3D 
        Private mDevice As Device
        Private mSurface As Surface
        Protected presentparams As PresentParameters
#End Region

#Region "MIEMBROS NECESARIOS PARA HACER FUNCIONAR A DIRECTSHOW"
        Private mFG As IFilterGraph2

        'INTERFACES PARA CONTROL DEL STREAM
        Private mAudio As IBasicAudio
        Private mVideo As IBasicVideo
        Private mControl As IMediaControl
        Private mEvents As IMediaEventEx
        Private mSeek As IMediaSeeking

        'VMR9 RENDERER INTERFACES
        Private mVMR9 As IBaseFilter
        Private mVMRWC As IVMRWindowlessControl9
        Private mVMRMixer As IVMRMixerControl9
        Private mVMRQualProp As IQualProp

        'BITMAP PARA LA IMAGEN CON EL TEXTO
        Private BMP As System.Drawing.Bitmap

        'SOURCE FILTER
        Private mSourceFilter As IBaseFilter

        'LAV SPLITTER
        Private mLAVSplitter As IBaseFilter

        'LAV VIDEO DECODER
        Private mLAVVideoDecoder As IBaseFilter

        'LAV AUDIO DECODER
        Private mLAVAudioDecoder As IBaseFilter

        'DIRECTSOUND DEFAULT DEVICE
        Private mDSoundRenderer As IBaseFilter
        Private mSoundStats As IAMAudioRendererStats
#End Region

#Region "CONSTRUCTORES"

        Public Sub New()
            Me.StreamState = ENUMS.ESTADO.ESTADO_CERRADO
            Me.mHWND = IntPtr.Zero
            Me.mHasText = False
            Me.mHasVideo = False
            Me.mHasAudio = False

            Me.mTextParams.DrawBorder = True
            Me.mTextParams.FontColor = Color.White
            Me.mTextParams.FontFace = "Arial"
            Me.mTextParams.OutlineSize = 4
            Me.mTextParams.Bold = True
            Me.mTextParams.FontSize = 18.0F
        End Sub

#End Region

#Region "PROPIEDADES"

        Public Property FileName() As String
            Get
                Return Me.mFileName
            End Get
            Set(value As String)
                If System.IO.File.Exists(value) Then
                    Me.mFileName = value
                Else
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_RUTAINCORRECTA)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Handle de la ventana que reproduce el video
        ''' </summary>
        ''' <value>IntPtr handle de la ventana</value>
        ''' <returns>IntPtr handle de la ventana</returns>
        ''' <remarks></remarks>
        Public Property VideoWindow() As IntPtr
            Get
                Return Me.mHWND
            End Get
            Set(value As IntPtr)
                If WIN32_API.IsWindow(value) Then
                    Me.mHWND = value
                Else
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VENTANAINEXISTENTE)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Handle de la ventana principal: usado por Direct3D 
        ''' </summary>
        ''' <value>IntPtr handle de la ventana</value>
        ''' <returns>IntPtr handle de la ventana</returns>
        ''' <remarks></remarks>
        Public Property MainHandle() As IntPtr
            Get
                Return Me.mMainHWND
            End Get
            Set(value As IntPtr)
                If WIN32_API.IsWindow(value) Then
                    Me.mMainHWND = value
                Else
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VENTANAINEXISTENTE)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Obtiene o establece el tamaño del video
        ''' </summary>
        ''' <value>Size: Nuevo tamaño</value>
        ''' <returns>Size: Tamaño actual del video</returns>
        ''' <remarks></remarks>
        Public Property VideoSize() As Size
            Get
                If Me.mVMRWC IsNot Nothing Then
                    Return Me.mVideoSize
                Else
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                End If
            End Get
            Set(value As Size)
                If Me.mVMRWC IsNot Nothing Then
                    Dim dsDst As New DsRect, dsSrc As New DsRect
                    Me.mVMRWC.GetVideoPosition(dsDst, dsSrc)
                    'dsDst.right = value.Width
                    'dsDst.bottom = value.Height
                    dsSrc.left = 0
                    dsSrc.top = 0
                    dsSrc.right = value.Width
                    dsSrc.bottom = value.Height

                    If Me.mVMRWC.SetVideoPosition(Nothing, dsSrc) = 0 Then
                        Me.mVideoSize = value
                    Else
                        RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    End If
                End If
            End Set
        End Property

        ''' <summary>
        ''' Obtiene el tamaño original del video
        ''' </summary>
        ''' <value></value>
        ''' <returns>Size: Tamaño del video</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property NativeVideoSize() As Size
            Get
                'SI NO HAY VIDEO O LAS INTERFACES NO FUNCAN VOLVEMOS INFORMANDO ERROR (NO EXISTE VIDEO DE 0x0
                If Me.mVMRWC Is Nothing OrElse (Not Me.mHasVideo) Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                    Return New Size(0, 0)
                End If

                Dim s As New Size
                If Me.mVMRWC.GetNativeVideoSize(s.Width, s.Height, New Integer, New Integer) <> 0 Then RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                Return s
            End Get
        End Property

        ''' <summary>
        ''' Determina si el stream posee video: solo funciona con CustomFG
        ''' </summary>
        ''' <value></value>
        ''' <returns>Boolean: True si hay video, False si no lo hay</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property HasVideo() As Boolean
            Get
                'NI SIQUIERA HAY STREAM ABIERTO
                If Me.StreamState = ENUMS.ESTADO.ESTADO_CERRADO Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                    Return False
                End If

                If Me.mLAVSplitter IsNot Nothing Then
                    'ESTAMOS USANDO NUESTRO GRAFO PERSONALIZADO, PODEMOS LLAMAR A cPlayerV2::CustomGraphHasVideo
                    Return Me.mHasVideo
                Else
                    'PROBABLEMENTE ESTEMOS USANDO INTELIGENT CONNECT, NO QUEDA OTRA QUE PREGUNTARLE A VMR9
                    If Me.mVMRWC Is Nothing Then
                        RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                        Return False
                    End If

                    Dim _w As Integer, _h As Integer, _arw As Integer, _arh As Integer
                    If Me.mVMRWC.GetNativeVideoSize(_w, _h, _arw, _arh) <> 0 Then RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)

                    'SI _W = 0 Y _H = 0 PROBABLEMENTE NO HAYA VIDEO
                    If _w = 0 AndAlso _h = 0 Then Return False
                    Return True
                End If
            End Get
        End Property

        ''' <summary>
        ''' Determina si el stream posee Audio: solo funciona con CustomFG
        ''' </summary>
        ''' <value></value>
        ''' <returns>Boolean: True si hay audio, False si no lo hay</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property HasAudio() As Boolean
            Get
                'NI SIQUIERA HAY UN STREAM ABIERTO
                If Me.mFG Is Nothing Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                    Return False
                End If

                'SI NO TENEMOS ACCESO A NUESTRO FILTRO, NO PODEMOS DETERMINARLO, DEVOLVEMOS TRUE X LAS DUDAS
                If Me.mLAVSplitter Is Nothing Then Return True
                Return Me.mHasAudio
            End Get
        End Property

        ''' <summary>
        ''' Determina si hay un texto dibujado sobre el stream actualmente
        ''' </summary>
        ''' <value></value>
        ''' <returns>Boolean: True si hay un texto dibujado, False si no lo hay</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property HasText() As Boolean
            Get
                Return Me.mHasText
            End Get
        End Property

        ''' <summary>
        ''' Obtiene o establece el formato de tiempo para el stream multimedia
        ''' Los posibles valores son:
        ''' Milliseconds
        ''' Seconds
        ''' Frames
        ''' Samples
        ''' </summary>
        ''' <value>ENUMS.TIME_FORMAT: miembro de la enumeracion que determina el formato actual</value>
        ''' <returns>ENUMS.TIME_FORMAT: formato de tiempo actual</returns>
        ''' <remarks></remarks>
        Public Property TimeFormat() As ENUMS.TIME_FORMAT
            Get
                Return Me.mTimeFormat
            End Get
            Set(value As ENUMS.TIME_FORMAT)
                Select Case value
                    Case ENUMS.TIME_FORMAT.TF_FRAMES
                        If Me.mSeek.IsFormatSupported(DirectShowLib.TimeFormat.Frame) = 0 Then
                            Me.mSeek.SetTimeFormat(DirectShowLib.TimeFormat.Frame)
                            Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_FRAMES
                            RaiseEvent TimeFormatChanged()
                        Else
                            RaiseEvent Error(ENUMS.ERROR_LIST.ERL_TFNOSOPORTADO)
                        End If
                    Case ENUMS.TIME_FORMAT.TF_SECONDS
                        If Me.mSeek.IsFormatSupported(DirectShowLib.TimeFormat.MediaTime) = 0 Then
                            Me.mSeek.SetTimeFormat(DirectShowLib.TimeFormat.MediaTime)
                            Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_SECONDS
                            RaiseEvent TimeFormatChanged()
                        Else
                            RaiseEvent Error(ENUMS.ERROR_LIST.ERL_TFNOSOPORTADO)
                        End If
                    Case ENUMS.TIME_FORMAT.TF_MILLISECONDS
                        If Me.mSeek.IsFormatSupported(DirectShowLib.TimeFormat.MediaTime) = 0 Then
                            Me.mSeek.SetTimeFormat(DirectShowLib.TimeFormat.MediaTime)
                            Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_MILLISECONDS
                            RaiseEvent TimeFormatChanged()
                        Else
                            RaiseEvent Error(ENUMS.ERROR_LIST.ERL_TFNOSOPORTADO)
                        End If
                    Case ENUMS.TIME_FORMAT.TF_SAMPLES
                        If Me.mSeek.IsFormatSupported(DirectShowLib.TimeFormat.Sample) = 0 Then
                            Me.mSeek.SetTimeFormat(DirectShowLib.TimeFormat.Sample)
                            Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_SAMPLES
                            RaiseEvent TimeFormatChanged()
                        Else
                            RaiseEvent Error(ENUMS.ERROR_LIST.ERL_TFNOSOPORTADO)
                        End If
                End Select
            End Set
        End Property

        ''' <summary>
        ''' Obtiene o establece la posicion actual dentro del stream multimedia
        ''' El mismo sera establecido utilizando el formato de tiempo actual
        ''' </summary>
        ''' <value>Long: Nueva posicion dentro del stream (respetando formato de tiempo actual)</value>
        ''' <returns>Long: Posicion actual dentro del stream (respetando formato de tiempo actual</returns>
        ''' <remarks></remarks>
        Public Property Position() As Long
            Get
                'EVITAMOS EXCEPCIONES
                If Me.mSeek Is Nothing Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                    Return -1
                End If

                'DETERMINAMOS LA POSICION ACTUAL
                Dim pos As Long
                'EN REALIDAD STOP POSITION NO NOS INTERESA YA QUE SIEMPRE LO HAREMOS AL FINAL DEL STREAM
                Me.mSeek.GetPositions(pos, New Long)

                If Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_MILLISECONDS Then
                    Return (pos / 10000)
                ElseIf Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_SECONDS Then
                    Return ((pos / 10000) / 1000)
                Else
                    Return pos
                End If
            End Get
            Set(value As Long)
                If value < 0 OrElse value > Me.Duration Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VALORINCORRECTO)
                    Exit Property
                End If

                Dim pos As Long

                If Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_MILLISECONDS Then
                    pos = value * 10000
                ElseIf Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_SECONDS Then
                    pos = value * 10000 * 1000
                Else
                    pos = value
                End If

                If Me.mSeek.SetPositions(New DirectShowLib.DsLong(value), AMSeekingSeekingFlags.AbsolutePositioning, New DsLong(Me.Duration), AMSeekingSeekingFlags.NoPositioning) <> 0 Then RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
            End Set
        End Property

        ''' <summary>
        ''' Solo Lectura: Obtiene la duracion del stream multimedia respetando formato de tiempo actual
        ''' </summary>
        ''' <value></value>
        ''' <returns>Long: Duracion del stream multimedia (respetando formato de tiempo actual)</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Duration() As Long
            Get
                'EVITAMOS EXCEPCIONES
                If Me.mSeek Is Nothing Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                    Return -1
                End If

                'DETERMINAMOS LA DURACION
                Dim mDuration As Long
                If Me.mSeek.GetDuration(mDuration) <> 0 Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    Return -1
                End If

                Select Case Me.mTimeFormat
                    Case ENUMS.TIME_FORMAT.TF_MILLISECONDS
                        mDuration = mDuration / 10000
                    Case ENUMS.TIME_FORMAT.TF_SECONDS
                        mDuration = mDuration / 10000 / 1000
                End Select
                Return mDuration
            End Get
        End Property

        ''' <summary>
        ''' Obtiene o establece la velocidad de reproduccion, los valores aceptados van desde 0.0F hasta 2.0F siendo 1.0F la velocidad estandar
        ''' </summary>
        ''' <value>Double: nueva PlayRate para el stream multimedia actual</value>
        ''' <returns>Double: PlayRate del stream multimedia actual</returns>
        ''' <remarks></remarks>
        Public Property Rate() As Double
            Get
                Dim dwRate As Double
                If Me.mSeek Is Nothing Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                    Return -1
                End If

                If Me.mSeek.GetRate(dwRate) <> 0 Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    Return -1
                End If

                Return dwRate
            End Get
            Set(value As Double)
                'VER EN MSDN PARA MAYOR INFORMACION 
                If value <= 0 OrElse value > 2 Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VALORINCORRECTO)
                    Exit Property
                End If

                If Me.mSeek.SetRate(value) <> 0 Then RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
            End Set
        End Property

        ''' <summary>
        ''' Obtiene o establece el balance actual (solo se aprecia cuando se trata de sonido Stereo)
        ''' </summary>
        ''' <value>Long: -100 para canal izquierdo, 100 para canal derecho, 0 Balance medio</value>
        ''' <returns>Long: Balance actual del stream multimedia</returns>
        ''' <remarks></remarks>
        Public Property Balance() As Long
            Get
                Dim mBal As Integer = 0
                If Not Me.mAudio Is Nothing Then
                    If Me.mAudio.get_Balance(mBal) <> 0 Then RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                End If
                Return mBal / 100
            End Get
            Set(ByVal value As Long)
                If value < -100 Or value > 100 Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VALORINCORRECTO)
                    Exit Property
                End If

                Dim mBal As Long = value * 100
                If Not Me.mAudio Is Nothing Then
                    If Me.mAudio.put_Balance(mBal) = 0 Then
                        RaiseEvent BalanceChanged()
                    Else
                        RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    End If

                End If
            End Set
        End Property

        ''' <summary>
        ''' Solo Lectura: Obtiene el estado del stream
        ''' </summary>
        ''' <value></value>
        ''' <returns>ENUMS.ESTADO: Posible valores
        ''' CERRADO
        ''' ABIERTO
        ''' REPRODUCIENDO
        ''' PAUSADO
        ''' PARADO</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property State() As ENUMS.ESTADO
            Get
                Return Me.StreamState
            End Get
        End Property

        ''' <summary>
        ''' Obtiene o establece el volumen del stream, notar que la funcion es exponencial 
        ''' por lo cual el valor en dB no siempre se corresopndera con el valor especificado
        ''' </summary>
        ''' <value>Long: Nuevo volumen del stream desde 0 hasta 100 donde 0 significa silencio y 100 maximo volumen (10000 dBs)</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Volume() As Long
            Get
                Dim mVol As Integer
                If Not Me.mAudio Is Nothing Then
                    If Me.mAudio.get_Volume(mVol) <> 0 Then RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                End If

                If mVol >= -1000 Or mVol <= 0 Then Return ((mVol / 100) + 100)
                Return -1
            End Get
            Set(ByVal value As Long)
                If value < 0 Or value > 100 Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VALORINCORRECTO)
                    Exit Property
                End If

                Dim mVol As Long = (100 - value) * (-100)
                If Not Me.mAudio Is Nothing Then
                    If Me.mAudio.put_Volume(mVol) = 0 Then
                        RaiseEvent VolumeChanged()
                    Else
                        RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    End If
                End If
            End Set
        End Property

#Region "PROPIEDADES RELATIVAS A TEXT OVERLAY"
        Public Property TextFontFace() As String
            Get
                Return Me.mTextParams.FontFace
            End Get
            Set(value As String)
                Me.mTextParams.FontFace = value
            End Set
        End Property

        Public Property TextFontSize() As Single
            Get
                Return Me.mTextParams.FontSize
            End Get
            Set(value As Single)
                If value < 0 Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VALORINCORRECTO)
                Else
                    Me.mTextParams.FontSize = value
                End If
            End Set
        End Property

        Public Property TextBold() As Boolean
            Get
                Return Me.mTextParams.Bold
            End Get
            Set(value As Boolean)
                Me.mTextParams.Bold = value
            End Set
        End Property

        Public Property TextItalic() As Boolean
            Get
                Return Me.mTextParams.Italic
            End Get
            Set(value As Boolean)
                Me.mTextParams.Italic = value
            End Set
        End Property

        Public Property TextUnderline() As Boolean
            Get
                Return Me.mTextParams.UnderLine
            End Get
            Set(value As Boolean)
                Me.mTextParams.UnderLine = value
            End Set
        End Property

        Public Property TextDrawBorder() As Boolean
            Get
                Return Me.mTextParams.DrawBorder
            End Get
            Set(value As Boolean)
                Me.mTextParams.DrawBorder = value
            End Set
        End Property

        Public Property TextFontColor() As Color
            Get
                Return Me.mTextParams.FontColor
            End Get
            Set(value As Color)
                Me.mTextParams.FontColor = value
            End Set
        End Property

        Public Property TextOutlineSize() As Single
            Get
                Return Me.mTextParams.OutlineSize
            End Get
            Set(value As Single)
                If value <= 0 Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VALORINCORRECTO)
                Else
                    Me.mTextParams.OutlineSize = value
                End If
            End Set
        End Property
#End Region

#End Region

#Region "METODOS PUBLICOS"
        ''' <summary>
        ''' (EXPERIMENTAL) : en caso de tratarse de video, permite recorrer el stream cuadro a cuadro (Frame Per Frame)
        ''' el stream debe soportar TIME_FORMAT.Frame
        ''' </summary>
        ''' <param name="uUnits">Cantidad de cuadros que se avanzaran dentro del stream</param>
        ''' <remarks></remarks>
        Public Sub FF(ByVal uUnits As Long)
            If Me.mSeek Is Nothing Then Exit Sub

            'SOLO SE ACEPTARA EN VIDEOS
            If Not Me.mHasVideo Then Exit Sub

            'MOMENTANEAMENTE ESTABLECEMOS EL TIEMPO A FRAMES 
            Dim tempTimeFormat As ENUMS.TIME_FORMAT
            tempTimeFormat = Me.mTimeFormat

            Me.TimeFormat = ENUMS.TIME_FORMAT.TF_FRAMES

            Me.mSeek.SetPositions(New DsLong(uUnits), AMSeekingSeekingFlags.RelativePositioning, 0, AMSeekingSeekingFlags.NoPositioning)

            Me.TimeFormat = tempTimeFormat
        End Sub

        ''' <summary>
        ''' (EXPERIMENTAL): en caso de tratarse de video, permite recorrer el stream cuadro a cuadro (Frame Per Frame)
        ''' el stream debe soportar TIME_FORMAT.Frame
        ''' </summary>
        ''' <param name="uUnits">Cantidad de cuadros que se atrasaran dentro del stream</param>
        ''' <remarks></remarks>
        Public Sub RW(ByVal uUnits As Long)
            If Me.mSeek Is Nothing Then Exit Sub

            'SOLO SE ACEPTARA EN VIDEOS
            If Not Me.mHasVideo Then Exit Sub

            'MOMENTANEAMENTE ESTABLECEMOS EL TIEMPO A FRAMES 
            Dim tempTimeFormat As ENUMS.TIME_FORMAT
            tempTimeFormat = Me.mTimeFormat

            Me.TimeFormat = ENUMS.TIME_FORMAT.TF_FRAMES

            Me.mSeek.SetPositions(New DsLong(-uUnits), AMSeekingSeekingFlags.RelativePositioning, 0, AMSeekingSeekingFlags.NoPositioning)

            Me.TimeFormat = tempTimeFormat
        End Sub

        ''' <summary>
        ''' Guarda en un archivo JPEG el cuadro actual de video
        ''' </summary>
        ''' <param name="sFileName">Ruta del archivo en el que se guardara la imagen</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCurrentImage(ByVal sFileName As String) As Boolean
            Dim BMP As System.Drawing.Bitmap = Nothing
            Dim ImgPtr As IntPtr = IntPtr.Zero
            Dim Header As New DirectShowLib.BitmapInfoHeader

            Try
                If Not Me.mVMRWC Is Nothing Then
                    DirectShowLib.DsError.ThrowExceptionForHR(Me.mVMRWC.GetCurrentImage(ImgPtr))
                    Marshal.PtrToStructure(ImgPtr, Header)

                    BMP = New Bitmap(Header.Width, Header.Height, (Header.BitCount / 8) * Header.Width, System.Drawing.Imaging.PixelFormat.Format32bppArgb, New IntPtr(ImgPtr.ToInt64 + 40))
                    BMP.RotateFlip(System.Drawing.RotateFlipType.RotateNoneFlipY)
                    BMP.Save(sFileName, System.Drawing.Imaging.ImageFormat.Jpeg)
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            Finally
                If Not BMP Is Nothing Then BMP.Dispose()
                If ImgPtr <> IntPtr.Zero Then Marshal.FreeCoTaskMem(ImgPtr)
            End Try
        End Function

#Region "VIDEO_PROCAMP_CONTROL"

        Public Function GetValues(ByVal sProp As ENUMS.PROP) As ESTRUCTURAS.M_M_V_D
            Dim vmr9ProcAmp As New DirectShowLib.VMR9ProcAmpControl
            Dim vmr9MixerControl As DirectShowLib.IVMRMixerControl9
            Dim vmr9ProcAmpRange As New DirectShowLib.VMR9ProcAmpControlRange
            Dim retVal As New ESTRUCTURAS.M_M_V_D

            If Me.mVMR9 Is Nothing Then Return Nothing

            vmr9MixerControl = DirectCast(Me.mVMR9, DirectShowLib.IVMRMixerControl9)

            vmr9ProcAmp.dwSize = Marshal.SizeOf(vmr9ProcAmp)


            If sProp = ENUMS.PROP.BRIGHNESS Then
                retVal.dwCurrrentValue = vmr9ProcAmp.Brightness
                vmr9ProcAmpRange.dwProperty = DirectShowLib.VMR9ProcAmpControlFlags.Brightness
                vmr9ProcAmp.dwFlags = DirectShowLib.VMR9ProcAmpControlFlags.Brightness
            ElseIf sProp = ENUMS.PROP.CONTRAST Then
                retVal.dwCurrrentValue = vmr9ProcAmp.Contrast
                vmr9ProcAmpRange.dwProperty = DirectShowLib.VMR9ProcAmpControlFlags.Contrast
                vmr9ProcAmp.dwFlags = DirectShowLib.VMR9ProcAmpControlFlags.Contrast
            ElseIf sProp = ENUMS.PROP.HUE Then
                retVal.dwCurrrentValue = vmr9ProcAmp.Hue
                vmr9ProcAmpRange.dwProperty = DirectShowLib.VMR9ProcAmpControlFlags.Hue
                vmr9ProcAmp.dwFlags = DirectShowLib.VMR9ProcAmpControlFlags.Hue
            ElseIf sProp = ENUMS.PROP.SATURATION Then
                retVal.dwCurrrentValue = vmr9ProcAmp.Saturation
                vmr9ProcAmpRange.dwProperty = DirectShowLib.VMR9ProcAmpControlFlags.Saturation
                vmr9ProcAmp.dwFlags = DirectShowLib.VMR9ProcAmpControlFlags.Saturation
            End If

            vmr9ProcAmpRange.dwSize = Marshal.SizeOf(vmr9ProcAmpRange)

            vmr9MixerControl.GetProcAmpControlRange(0, vmr9ProcAmpRange)

            vmr9MixerControl.GetProcAmpControl(0, vmr9ProcAmp)

            retVal.dwMinValue = vmr9ProcAmpRange.MinValue
            retVal.dwMaxValue = vmr9ProcAmpRange.MaxValue
            retVal.dwStep = vmr9ProcAmpRange.StepSize
            retVal.dwDefault = vmr9ProcAmpRange.DefaultValue

            Return retVal
        End Function

        Public Function SetProcAmpValue(ByVal PropToSet As ENUMS.PROP, ByVal dwVal As Single) As Boolean
            Try
                If Me.mVMR9 Is Nothing Then Return False

                Dim allowedVals As ESTRUCTURAS.M_M_V_D = Me.GetValues(PropToSet)

                If dwVal > allowedVals.dwMaxValue Or dwVal < allowedVals.dwMinValue Then Throw New IndexOutOfRangeException

                Dim VMR9ProcAmp As New DirectShowLib.VMR9ProcAmpControl

                VMR9ProcAmp.dwSize = Marshal.SizeOf(VMR9ProcAmp)

                Select Case PropToSet
                    Case ENUMS.PROP.BRIGHNESS
                        VMR9ProcAmp.dwFlags = DirectShowLib.VMR9ProcAmpControlFlags.Brightness
                        VMR9ProcAmp.Brightness = dwVal
                    Case ENUMS.PROP.CONTRAST
                        VMR9ProcAmp.Contrast = CSng(allowedVals.dwMinValue + dwVal * (allowedVals.dwMaxValue - allowedVals.dwMinValue))
                    Case ENUMS.PROP.HUE
                        VMR9ProcAmp.dwFlags = DirectShowLib.VMR9ProcAmpControlFlags.Hue
                        VMR9ProcAmp.Hue = dwVal
                    Case ENUMS.PROP.SATURATION
                        VMR9ProcAmp.dwFlags = DirectShowLib.VMR9ProcAmpControlFlags.Saturation
                        VMR9ProcAmp.Saturation = dwVal
                End Select

                Dim VMRProcAmp As DirectShowLib.IVMRMixerControl9 = DirectCast(Me.mVMR9, DirectShowLib.IVMRMixerControl9)

                DirectShowLib.DsError.ThrowExceptionForHR(VMRProcAmp.SetProcAmpControl(0, VMR9ProcAmp))

                Return True

            Catch ex As Exception
                Return False
            End Try
        End Function

#End Region

        ''' <summary>
        ''' Obtiene el cuadro actual de video siendo mostrado
        ''' </summary>
        ''' <param name="retBMP">Objeto de tipo Image en el cual se depositara el cuadro actual</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCurrentImage(ByRef retBMP As Image) As Boolean
            Dim BMP As System.Drawing.Bitmap = Nothing
            Dim ImgPtr As IntPtr = IntPtr.Zero
            Dim Header As New DirectShowLib.BitmapInfoHeader

            Try
                If Not Me.mVMRWC Is Nothing Then
                    DirectShowLib.DsError.ThrowExceptionForHR(Me.mVMRWC.GetCurrentImage(ImgPtr))
                    Marshal.PtrToStructure(ImgPtr, Header)

                    BMP = New Bitmap(Header.Width, Header.Height, (Header.BitCount / 8) * Header.Width, System.Drawing.Imaging.PixelFormat.Format32bppArgb, New IntPtr(ImgPtr.ToInt64 + 40))
                    BMP.RotateFlip(System.Drawing.RotateFlipType.RotateNoneFlipY)

                    retBMP = BMP
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            Finally
                If Not BMP Is Nothing Then BMP.Dispose()
                If ImgPtr <> IntPtr.Zero Then Marshal.FreeCoTaskMem(ImgPtr)
            End Try
        End Function

        ''' <summary>
        ''' Abre un stream multimedia, por orden la operacion se realiza de la siguiente manera:
        ''' Se intenta armar el CustomFilterGraph (usando el set de filtros LAV)
        ''' Se intenta armar el FilterGraph utilizando IntelligentConnect
        ''' Falla la renderizacion del archivo multimedia
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Open() As Boolean
            If Me.mFG IsNot Nothing Then Me.CloseAndRelease()
            Me.mFG = DirectCast(New FilterGraph, IFilterGraph2)

            If Me.mFileName = "" Then Return False

            'INTENTAMOS RENDERIZAR CON EL FG CUSTOMIZADO
            If Me.CreateCustomFilterGraph(Me.mFileName) = False Then
                'NO SE PUDO, LIMPIAMOS LO QUE EL FG CUSTOMIZADO CREO Y VAMOS DE NUEVO CON INTELIGENT CONNECT
                Me.Close()
                Me.mFG = DirectCast(New FilterGraph, IFilterGraph2)
                If Me.RenderIntelligentConnect(Me.mFileName) = False Then
                    'OOPS TAMPOCO SE PUDO, COMO BIEN DIJIMOS, TAMPOCO SOMOS MAGOS, SI NO VA, NO VA
                    Me.StreamState = ENUMS.ESTADO.ESTADO_CERRADO
                    Return False
                End If
            End If

            'OKEY, SE PUDO CREAR EL FG Y RENDERIZAR EL ARCHIVO MM, PROCEDEMOS A 
            'APLICAR LOS METODOS QueryInterface SOBRE EL FILTER GRAPH
            Me.mControl = DirectCast(Me.mFG, IMediaControl)
            Me.mAudio = DirectCast(Me.mFG, IBasicAudio)
            Me.mSeek = DirectCast(Me.mFG, IMediaSeeking)

            'OBTENEMOS EL FORMATO ACTUAL DE TIEMPO Y LO CARGAMOS EN LA VARIABLE DE INSTANCIA
            Dim g As New Guid

            Me.mSeek.GetTimeFormat(g)

            Select Case g
                Case DirectShowLib.TimeFormat.MediaTime
                    'haremos las divisiones correspondientes luego
                    Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_MILLISECONDS
                Case DirectShowLib.TimeFormat.Frame
                    Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_FRAMES
                Case DirectShowLib.TimeFormat.Sample
                    Me.mTimeFormat = ENUMS.TIME_FORMAT.TF_SAMPLES
            End Select

            Me.StreamState = ENUMS.ESTADO.ESTADO_ABIERTO
            If Me.HasVideo Then Me.VideoSize = Me.NativeVideoSize
            RaiseEvent FileOpened()
            Return True
        End Function

        ''' <summary>
        ''' Frena la reproduccion y libera todos los recursos utilizados por los objetos no administrados
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Close() As Boolean
            Me.ReleaseDirect3D()
            Return Me.CloseAndRelease()
            Me.mHasVideo = False
            Me.mHasAudio = False
            Me.mHasText = False
        End Function

        ''' <summary>
        ''' Comienza la reproduccion del stream abierto actualmente
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Play() As Boolean
            If Me.mControl Is Nothing Then Return False
            Me.mControl.Run()
            Me.StreamState = ENUMS.ESTADO.ESTADO_REPRODUCIENDO
            Return True
        End Function

        ''' <summary>
        ''' Pausa la reproduccion del stream abierto actualmente
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Pause() As Boolean
            If Me.mControl Is Nothing Then Return False
            Me.mControl.Pause()
            Me.StreamState = ENUMS.ESTADO.ESTADO_PAUSADO
            Return True
        End Function

        ''' <summary>
        ''' Frena la reproduccion del stream abierto actualmente y establece la posicion en 0
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function [Stop]() As Boolean
            If Me.mControl Is Nothing Then Return False
            Me.mControl.Stop()
            Me.Position = 0
            Me.StreamState = ENUMS.ESTADO.ESTADO_DETENIDO
            Return True
        End Function

        ''' <summary>
        ''' Indica a VMR9 que se debe repintar el cuadro de video, este metodo deberia llamarse en el 
        ''' evento Paint del control en el cual se reproduce el video
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub RepaintVideo()
            If Me.mVMRWC Is Nothing Then Exit Sub
            If Me.mHWND = IntPtr.Zero Then Exit Sub

            Me.mVMRWC.RepaintVideo(Me.mHWND, WIN32_API.GetDC(Me.mHWND))
        End Sub

        ''' <summary>
        ''' Muestra u oculta un texto sobre el stream de video
        ''' </summary>
        ''' <param name="szText">Texto que se dibujara, si bShow es igual a False este argumento puede ser nulo</param>
        ''' <param name="bShow">Determina si se debe mostrar o eliminar el texto del stream de video</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function WriteOverlayText(ByVal szText As String, ByVal bShow As Boolean) As Boolean
            If Me.mVMR9 Is Nothing Then Return False
            If Me.mHasVideo = False Then Return False

            Dim vmrBMP As IVMRMixerBitmap9
            Dim vmrparams As New VMR9AlphaBitmap
            Dim hImg As Image

            'If Me.BMP IsNot Nothing Then Me.BMP.Dispose()
            'Me.BMP = Nothing

            Try
                vmrBMP = DirectCast(Me.mVMR9, IVMRMixerBitmap9)

                DsError.ThrowExceptionForHR(vmrBMP.GetAlphaBitmapParameters(vmrparams))

                If bShow Then
                    Me.BMP = Me.CreateRectText(szText, New Size(Me.mVideoSize.Width, Me.mVideoSize.Height), New System.Drawing.FontFamily(Me.mTextParams.FontFace), Me.mTextParams.FontColor, Me.mTextParams.AlphaColor, Me.mTextParams.DrawBorder, Me.mTextParams.UnderLine, Me.mTextParams.Bold, Me.mTextParams.Italic, False, Me.mTextParams.OutlineSize, Me.mTextParams.FontSize)

                    If Not Me.InitializeDirect3D() Then Return False

                    'vmrparams.clrSrcKey = ColorTranslator.ToWin32(Me.mTextParams.AlphaColor)
                    vmrparams.dwFlags = VMR9AlphaBitmapFlags.EntireDDS

                    vmrparams.fAlpha = 1.0F
                    vmrparams.pDDS = Me.mSurface.GetObjectByValue(Me.DxMagicNumber)
                    vmrparams.rDest = New NormalizedRect(0.0F, 0.0F, 1.0F, 1.0F)

                    DsError.ThrowExceptionForHR(vmrBMP.SetAlphaBitmap(vmrparams))
                    Me.mHasText = True
                Else
                    vmrparams.dwFlags = VMR9AlphaBitmapFlags.Disable
                    DsError.ThrowExceptionForHR(vmrBMP.UpdateAlphaBitmapParameters(vmrparams))
                    Me.mHasText = False
                End If

                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function GetVideoStats(ByRef dwStats As ESTRUCTURAS.VIDEO_STATS) As Boolean
            Try
                If Me.mVMR9 Is Nothing Then Return False

                Me.mVMRQualProp = DirectCast(Me.mVMR9, IQualProp)

                If Me.mVMRQualProp Is Nothing Then Return False

                DsError.ThrowExceptionForHR(Me.mVMRQualProp.get_AvgFrameRate(dwStats.AVG_FRAME_RATE))
                DsError.ThrowExceptionForHR(Me.mVMRQualProp.get_FramesDrawn(dwStats.FRAMES_DRAWN))
                DsError.ThrowExceptionForHR(Me.mVMRQualProp.get_AvgSyncOffset(dwStats.AVG_SYNC_OFFSET))
                DsError.ThrowExceptionForHR(Me.mVMRQualProp.get_FramesDroppedInRenderer(dwStats.FRAMES_DROPPED))

                Me.mVMRQualProp = Nothing
                Return True

            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function GetAudioInfo(ByRef dwRet As ESTRUCTURAS.AUDIO_STATS) As Boolean
            If Me.mLAVAudioDecoder Is Nothing Then Return False

            If Me.StreamState = ENUMS.ESTADO.ESTADO_CERRADO Then
                'LA INFORMACION NO SERA LA CORRECTA 
                Return False
            End If

            'OBTENEMOS EL MEDIA TYPE PARA EL RENDERER DE DIRECTSOUND
            Dim mt(0) As AMMediaType
            Dim pin(0) As IPin
            Dim pinEnum As IEnumPins
            Dim typeEnum As IEnumMediaTypes
            Dim audioFormat As WaveFormatEx

            Try
                Me.mSoundStats = DirectCast(Me.mDSoundRenderer, IAMAudioRendererStats)

                If Me.mSoundStats Is Nothing Then Return False

                'TRAEMOS EL UNICO PIN DEL RENDERER (INPUT PIN)
                Me.mLAVAudioDecoder.EnumPins(pinEnum)

                While pinEnum.Next(1, pin, IntPtr.Zero) = 0

                    'Y OBTENEMOS EL MEDIA TYPE PARA DICHO PIN
                    pin(0).EnumMediaTypes(typeEnum)

                    While typeEnum.Next(1, mt, IntPtr.Zero) = 0
                        If mt(0).subType = MediaSubType.PCM Then
                            'LO TENEMOS CASTEAMOS A WAVEFORMATEX

                            audioFormat = System.Runtime.InteropServices.Marshal.PtrToStructure(mt(0).formatPtr, GetType(WaveFormatEx))
                            DsUtils.FreeAMMediaType(mt(0))
                            mt(0) = Nothing
                        Else
                            'NO SIRVE, SOLO LIBERAMOS RAM Y SEGUIMOS
                            DsUtils.FreeAMMediaType(mt(0))
                            mt(0) = Nothing
                        End If
                    End While

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(typeEnum)
                    typeEnum = Nothing

                    If audioFormat IsNot Nothing Then
                        dwRet.AVG_BYTES_PER_SEC = audioFormat.nAvgBytesPerSec
                        dwRet.CHANNELS = audioFormat.nChannels
                        dwRet.SAMPLES_PER_SEC = audioFormat.nSamplesPerSec
                        dwRet.AVG_BYTES_PER_SEC = audioFormat.nAvgBytesPerSec
                        Me.mSoundStats.GetStatParam(AMAudioRendererStatParam.BufferFullness, dwRet.BUFFER_FILL, Nothing)
                        Me.mSoundStats.GetStatParam(AMAudioRendererStatParam.LastBufferDur, dwRet.BUFFER_DURATION, Nothing)
                    End If

                    Marshal.ReleaseComObject(pin(0))
                End While

                System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                pinEnum = Nothing

                Return True
            Catch ex As Exception
                Return False
            Finally
                If pin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pin(0))
                If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                If typeEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(typeEnum)
                If mt(0) IsNot Nothing Then DsUtils.FreeAMMediaType(mt(0))
            End Try
        End Function

        ''' <summary>
        ''' Provee un metodo para determinar que esta sucediendo dentro del Filter Graph no muy util para
        ''' el usuario final pero puede proveer informacion importante para el desarrollador
        ''' </summary>
        ''' <param name="dwRet">Aqui se depositara la informacion obtenida del Filter Graph</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UnderTheHood(ByRef dwRet() As ESTRUCTURAS.UNDER_THE_HOOD) As Boolean
            If Me.mFG Is Nothing Then Return False

            Dim filterEnum As IEnumFilters
            Dim pinEnum As IEnumPins
            Dim mediaEnum As IEnumMediaTypes
            Dim i As Long = 0 'contador que se utilizara para contabilizar la cantidad de filtros
            Dim pinCount As Long = 0    'contador que se utilizara para contabilizar los pines por cada filtro
            Dim iFilter(0) As IBaseFilter 'aqui se depositaran los filtros
            Dim iPins(0) As IPin 'aqui se depositaran los pines
            Dim iMedia(0) As AMMediaType 'aqui se depositara informacion sobre tipos y subtipos soportados por cada pin

            Dim filterInfo As FilterInfo
            Dim PinInfo As PinInfo

            Try
                'OBTENEMOS UN PUNTERO A IEnumFilters para realizar la enumeracion
                Me.mFG.EnumFilters(filterEnum)
                ReDim dwRet(0)
                ReDim dwRet(0).Pins(0)

                While filterEnum.Next(1, iFilter, 0) = 0
                    'CARGAMOS NOMBRE DE FILTRO
                    iFilter(0).QueryFilterInfo(filterInfo)

                    dwRet(i).FilterName = filterInfo.achName

                    'LIBERAMOS RECURSOS
                    Marshal.ReleaseComObject(filterInfo.pGraph)

                    'ENUMERAMOS AHORA LOS PINES PARA ESTE FILTRO

                    iFilter(0).EnumPins(pinEnum)

                    While pinEnum.Next(1, iPins, 0) = 0
                        iPins(0).QueryPinInfo(PinInfo)

                        dwRet(i).Pins(pinCount).PinName = PinInfo.name
                        dwRet(i).Pins(pinCount).PinDirection = PinInfo.dir

                        'LIBERAMOS PININFO PORQUE YA NO LO UTILIZAREMOS
                        DsUtils.FreePinInfo(PinInfo)

                        'TRAEMOS EL MEDIATYPE PARA ESTE PIN
                        iPins(0).EnumMediaTypes(mediaEnum)

                        'LLENAMOS EL MEDIATYPE PARA ESTE PIN
                        mediaEnum.Next(1, iMedia, IntPtr.Zero)

                        dwRet(i).Pins(pinCount).MajorType = iMedia(0).majorType
                        dwRet(i).Pins(pinCount).SubType = iMedia(0).subType
                        dwRet(i).Pins(pinCount).SampleSize = iMedia(0).sampleSize

                        'LIBERAMOS MEMORIA USADA POR IMEDIA
                        DsUtils.FreeAMMediaType(iMedia(0))

                        'LIBERAMOS MEDIAENUM
                        Marshal.ReleaseComObject(mediaEnum)

                        'LIBERAMOS EL PIN
                        Marshal.ReleaseComObject(iPins(0))
                        ReDim Preserve dwRet(i).Pins(dwRet(i).Pins.Count)
                        pinCount += 1
                    End While
                    'LIBERAMOS PINENUM
                    Marshal.ReleaseComObject(pinEnum)

                    'LIBERAMOS IBASEFILTER
                    Marshal.ReleaseComObject(iFilter(0))
                    ReDim Preserve dwRet(dwRet.Count)
                    i += 1
                End While

                'LIBERAMOS FILTERENUM
                Marshal.ReleaseComObject(filterEnum)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

#End Region

#Region "METODOS PRIVADOS"
        ''' <summary>
        ''' Inicializa los objetos de Direct3D
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function InitializeDirect3D() As Boolean
            Device.IsUsingEventHandlers = False

            If Me.mMainHWND = IntPtr.Zero Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_VALORINCORRECTO)
                Return False
            End If

            'If Me.BMP Is Nothing Then Return False
            If Me.mVMRWC Is Nothing Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                Return False
            End If

            Me.presentparams = New PresentParameters
            Me.presentparams.Windowed = True
            Me.presentparams.SwapEffect = SwapEffect.Discard

            Me.mDevice = New Device(0, DeviceType.Hardware, Me.mMainHWND, CreateFlags.MultiThreaded Or CreateFlags.SoftwareVertexProcessing, Me.presentparams)

            If Me.mDevice Is Nothing Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_D3DDEVICE)
                Return False
            End If

            Me.mSurface = New Surface(Me.mDevice, Me.BMP, Pool.SystemMemory)

            If Me.mSurface Is Nothing Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_D3DSURFACE)
                Return False
            End If

            Return True
        End Function

        ''' <summary>
        ''' Libera los recursos utilizados por Direct3D 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ReleaseDirect3D() As Boolean
            Try
                If Me.mSurface IsNot Nothing Then
                    Me.mSurface.ReleaseGraphics()
                    Me.mSurface.Dispose()
                End If
                Me.mSurface = Nothing

                If Me.mDevice IsNot Nothing Then Me.mDevice.Dispose()
                Me.mDevice = Nothing
                If Me.BMP IsNot Nothing Then Me.BMP.Dispose()
                Me.BMP = Nothing
                Return True
            Catch ex As Exception
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                Return False
            End Try
        End Function

        Private Function CreateCustomFilterGraph(ByVal szFileName As String) As Boolean
            Dim bret As Boolean = True

            'AGREGAMOS EL SOURCEFILTER
            If Me.AddCustomFilter(True, "E436EBB5-524F-11CE-9F53-0020AF0BA770", "File Source (Async.)", Me.mSourceFilter, szFileName) = False Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                Return False
            End If

            'AGREGAMOS EL SPLITER
            If Me.AddCustomFilter(False, "171252A0-8820-4AFE-9DF8-5C92B2D66B04", "LAV Splitter", Me.mLAVSplitter) = False Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                Return False
            End If

            'CONECTAMOS LOS PINES DEL SOURCE FILTER CON EL SPLITTER
            bret = Me.ConnectCustomPinByNumber(Me.mSourceFilter, Me.mLAVSplitter, 0, 0, False)

            'AHORA VERIFICAMOS SI HAY AUDIO EN EL SOURCE, EN CASO DE HABERLO, CONECTAMOS EL DECODER DE AUDIO

            If Me.CustomGraphHasAudio(Me.mLAVSplitter) Then
                'AGREGAMOS EL AUDIO DECODER
                If Me.AddCustomFilter(False, "E8E73B6B-4CB3-44A4-BE99-4F7BCB96E491", "LAV Audio Decoder", Me.mLAVAudioDecoder) = False Then
                    Me.ReleaseCustomFG()
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    Return False
                End If

                'CONECTAMOS EL PIN 1 (AUDIO) DEL SPLITTER AL DECODER DE AUDIO
                If Me.ConnectCustomPinByName(Me.mLAVSplitter, Me.mLAVAudioDecoder, "Audio", "Input", False) = False Then
                    'LIMPIAMOS TODO Y VOLVEMOS
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    Me.ReleaseCustomFG()
                    Return False
                End If

                'AGREGAMOS EL AUDIO RENDERER (DEFAULT DSOUND RENDERER, X DONDE SALDRA ELSONIDO DEPENDERA DE LA CONFIGURACION DE SISTEMA)
                If Me.AddCustomFilter(False, "79376820-07D0-11CF-A24D-0020AFD79767", "Default DirectSound Device", Me.mDSoundRenderer, "") = False Then Return False

                'CONECTAMOS EL PIN 0 DEL DECODER DE AUDIO AL RENDERER DE AUDIO
                If Me.ConnectCustomPinByNumber(Me.mLAVAudioDecoder, Me.mDSoundRenderer, 0, 0, False) = False Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    Me.ReleaseCustomFG()
                    Return False
                End If
            End If

            If Me.CustomGraphHasVideo(Me.mLAVSplitter) Then
                'AGREGAMOS EL VIDEO DECODER
                If Me.AddCustomFilter(False, "EE30215D-164F-4A92-A4EB-9D4C13390F9F", "LAV Video Decoder", Me.mLAVVideoDecoder) = False Then Return False

                'CONECTAMOS EL PIN 0 (VIDEO) DEL SPLITTER AL DECODER DE VIDEO
                If Me.ConnectCustomPinByName(Me.mLAVSplitter, Me.mLAVVideoDecoder, "Video", "Input", False) = False Then
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    Me.ReleaseCustomFG()
                    Return False
                End If

                'AGREGAMOS EL VIDEO RENDERER (VIDEO MIXING RENDERER 9 EN ESTE CASO)
                If Me.AddCustomFilter(False, "51B4ABF3-748F-4E3B-A276-C828330E926A", "Video Mixing Renderer 9", Me.mVMR9) = False Then
                    Me.ReleaseCustomFG()
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    Return False
                End If

                'CONFIGURAMOS EL RENDERER
                bret = Me.ConfigureRenderer()

                'CONECTAMOS EL PIN 0 DEL DECODER DE VIDEO AL RENDERER DE VIDEO
                If Me.ConnectCustomPinByNumber(Me.mLAVVideoDecoder, Me.mVMR9, 0, 0, False) = False Then
                    Me.ReleaseCustomFG()
                    RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                    Return False
                End If
            End If

            'YA ESTA TODO CONECTADO, SI NO NOS MANDAMOS NINGUN PEDO AL DARLE A IMediaControl::Run() DEBERIA
            'COMENZAR AL REPRODUCCION CASO CONTRARIO, A LLORAR AL CAMPITO :P
            RaiseEvent CustomGraphRendered()
            Return True
        End Function

        Private Function ConfigureRenderer() As Boolean
            If Me.mVMR9 Is Nothing Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                Return False
            End If

            Dim mVMR9Config As IVMRFilterConfig9
            Try
                'OBTENEMOS LA INTERFACE DE CONFIGURACION DEL FILTRO DE VMR9
                mVMR9Config = DirectCast(Me.mVMR9, IVMRFilterConfig9)

                'CONFIGURAMOS CANTIDAD DE STREAMS (DE ESTE MODO VMR9 CARGARA EL MIXER EL CUAL UTILIZAREMOS LUEGO
                'Y TAMBIEN CONFIGURAMIOS MODO (WINDOWLESS PARA PROVEER NOSOTROS LA VENTANA DE REPRODUCCION)
                mVMR9Config.SetRenderingMode(VMR9Mode.Windowless)
                mVMR9Config.SetNumberOfStreams(1)

                'UNA VEZ CONFIGURADO PODEMOS TRAER LA INTERFACE PARA REPRODUCCION WINDOWLESS
                Me.mVMRWC = DirectCast(Me.mVMR9, IVMRWindowlessControl9)

                'CONFIGURAMOS LA VENTANA DE REPRODUCCION
                Me.mVMRWC.SetVideoClippingWindow(Me.mHWND)
                'Y EL ASPECTRATIO, EN PRINCIPIO LO VAMOS A PRESERVAR A NO SER QUE SE ESPECIFIQUE LO CONTRARIO
                Me.mVMRWC.SetAspectRatioMode(VMR9AspectRatioMode.LetterBox)
                Return True
            Catch ex As Exception
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                Return False
            End Try
        End Function

        Private Function CustomGraphHasAudio(ByVal splitterFilter As IBaseFilter) As Boolean
            If splitterFilter Is Nothing Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                Return False
            End If

            Try
                Dim pin(0) As IPin
                Dim pindir As PinDirection
                Dim pinEnum As IEnumPins
                Dim mt(0) As AMMediaType
                Dim mtenum As IEnumMediaTypes
                Dim bHasAudio As Boolean = False

                splitterFilter.EnumPins(pinEnum)

                While pinEnum.Next(1, pin, IntPtr.Zero) = 0
                    pin(0).QueryDirection(pindir)

                    If pindir = PinDirection.Output Then
                        'OBTENDREMOS EL MEDIA TYPE PARA ESTE PIN PARA VER SI ES VIDEO
                        pin(0).EnumMediaTypes(mtenum)
                        While mtenum.Next(1, mt, IntPtr.Zero) = 0
                            'VERIFICAMOS SI ALGUNO TIENE TIPO VIDEO ES SUF HAY VIDEO EN EL GRAFO
                            If mt(0).majorType = MediaType.Audio Or mt(0).majorType = MediaType.AnalogAudio Then
                                bHasAudio = True
                            End If
                            'LIBERAMOS AMMEDIATYPE
                            DsUtils.FreeAMMediaType(mt(0))
                        End While
                        'LIBERAMOS MTENUM
                        Marshal.ReleaseComObject(mtenum)
                    End If
                    'LIBERAMOS PIN
                    Marshal.ReleaseComObject(pin(0))
                    'SEGUIMOS CON LA ENUMERACION
                End While

                'LIBERAMOS PINENUM
                Marshal.ReleaseComObject(pinEnum)
                Me.mHasAudio = bHasAudio

                Return bHasAudio

            Catch ex As Exception
                Me.mHasAudio = False
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                Return False
            End Try
        End Function

        Private Function CustomGraphHasVideo(ByVal splitterFilter As IBaseFilter) As Boolean
            If splitterFilter Is Nothing Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                Return False
            End If

            Try
                Dim pin(0) As IPin
                Dim pindir As PinDirection
                Dim pinEnum As IEnumPins
                Dim mt(0) As AMMediaType
                Dim mtenum As IEnumMediaTypes
                Dim bHasVideo As Boolean = False

                splitterFilter.EnumPins(pinEnum)

                While pinEnum.Next(1, pin, IntPtr.Zero) = 0
                    pin(0).QueryDirection(pindir)

                    If pindir = PinDirection.Output Then
                        'OBTENDREMOS EL MEDIA TYPE PARA ESTE PIN PARA VER SI ES VIDEO
                        pin(0).EnumMediaTypes(mtenum)
                        While mtenum.Next(1, mt, IntPtr.Zero) = 0
                            'VERIFICAMOS SI ALGUNO TIENE TIPO VIDEO ES SUF HAY VIDEO EN EL GRAFO
                            If mt(0).majorType = MediaType.Video Then
                                bHasVideo = True
                            End If
                            'LIBERAMOS AMMEDIATYPE
                            DsUtils.FreeAMMediaType(mt(0))
                        End While
                        'LIBERAMOS MTENUM
                        Marshal.ReleaseComObject(mtenum)
                    End If
                    'LIBERAMOS PIN
                    Marshal.ReleaseComObject(pin(0))
                    'SEGUIMOS CON LA ENUMERACION
                End While

                'LIBERAMOS PINENUM
                Marshal.ReleaseComObject(pinEnum)
                Me.mHasVideo = bHasVideo
                Return bHasVideo

            Catch ex As Exception
                Me.mHasVideo = False
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                Return False
            End Try
        End Function

        Private Function ConnectCustomPinByNumber(ByVal USFilter As IBaseFilter, ByVal DSFilter As IBaseFilter, Optional ByVal USPinNumber As Integer = 0, Optional ByVal DSPinNumber As Integer = 0, Optional ByVal bUseICIfFailure As Boolean = False) As Boolean
            Dim bRet As Boolean = False
            Dim bFoundPin As Boolean = False

            Dim outputPin(0) As IPin 'OUTPUT PIN DEL USFILTER
            Dim inputPin(0) As IPin 'INPUT PIN DEL DSFILTER

            'OBTENEMOS EL PIN DE SALIDA DEL UPSTREAM FILTER
            Dim pdir As PinDirection
            Dim pinEnum As IEnumPins = Nothing

            Dim pinCount As Integer = 0

            Try
                USFilter.EnumPins(pinEnum)

                While pinEnum.Next(1, outputPin, IntPtr.Zero) = 0
                    outputPin(0).QueryDirection(pdir)
                    If pdir = PinDirection.Output Then
                        'TENEMOS EL PIN QUE DESEAMOS, VERIFICAMOS SI ES EL QUE BUSCAMOS POR INDICE
                        If pinCount = USPinNumber Then
                            'OK ESTE ES NUESTRO PIN
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                            pinEnum = Nothing
                            bFoundPin = True
                            Exit While
                        Else
                            'NO ES EL PIN QUE BUSCAMOS, SEGUIMOS LA BUSQUEDA
                            pinCount += 1
                        End If
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outputPin(0))
                End While

                If Not bFoundPin Then
                    'NO ENCONTRAMOS EL PIN QUE BUSCABAMOS, ABANDONAMOS
                    If outputPin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(outputPin(0))
                    If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                    Return False
                End If

                'NOS PREPARAMOS PARA LA SIGUIENTE ENUMERACION
                If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                pinCount = 0
                bFoundPin = False
                DSFilter.EnumPins(pinEnum)

                While pinEnum.Next(1, inputPin, IntPtr.Zero) = 0
                    inputPin(0).QueryDirection(pdir)

                    If pdir = PinDirection.Input Then
                        'TENEMOS UN PIN DE ENTRADA, VERIFICAMOS SI ES EL INDICE CORRECTO
                        If pinCount = DSPinNumber Then
                            'ESTE ES NUESTRO PIN
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                            pinEnum = Nothing
                            bFoundPin = True
                            Exit While
                        Else
                            'SEGUIMOS CON LA ENUMERACION
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(outputPin(0))
                        End If
                    End If
                End While

                If Not bFoundPin Then
                    'NO ENCONTRAMOS EL PIN QUE BUSCABAMOS, ABANDONAMOS
                    If inputPin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(inputPin(0))
                    If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                    Return False
                End If

                If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)

                'YA TENEMOS AMBOS PINES, INTENTAMOS CONECTARLOS DIRECTAMENTE
                'PASANDO NOTHING AL MEDIA TYPE SE DEJA QUE LOS FILTROS NEGOCIEN LA CONEXION ENTRE SI
                If Me.mFG.ConnectDirect(outputPin(0), inputPin(0), Nothing) <> 0 Then
                    'NO SE PUEDEN CONECTAR AMBOS PINES (NO SE PUEDE NEGOCIAR MEDIATYPE?)
                    If bUseICIfFailure Then
                        'EL DEV DESEA UTILIZAR INTELLIGENT CONNECT EN CASO DE FALLA, LO INTENTAMOS
                        If Me.mFG.Connect(outputPin(0), inputPin(0)) = 0 Then
                            'NO SE PUDO REALIZAR LA CONEXION
                            bRet = False
                        Else
                            'SE REALIZO LA CONEXION MEDIANTE IC
                            bRet = True
                        End If
                    Else
                        'EL USUARIO NO DESEA USAR INTELLIGENT CONNECT PERO LA CONEXION DIRECTA FALLO
                        bRet = False
                    End If
                Else
                    'LA CONEXION SE REALIZO DE UNA SIN USAR INTELLIGENT CONNECT
                    bRet = True
                End If

                'LIBERAMOS LAS INTERFACES DE LOS PINES
                If inputPin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(inputPin(0))
                If outputPin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(outputPin(0))
                If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)

                Return bRet

            Catch ex As Exception
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                Return False
            End Try
        End Function

        Private Function ConnectCustomPinByName(ByVal USFilter As IBaseFilter, ByVal DSFilter As IBaseFilter, ByVal USPinName As String, ByVal DSPinName As String, Optional ByVal bUseICIfFailure As Boolean = False) As Boolean
            Dim bRet As Boolean = False
            Dim bFoundPin As Boolean = False

            Dim outputPin(0) As IPin 'OUTPUT PIN DEL USFILTER
            Dim inputPin(0) As IPin 'INPUT PIN DEL DSFILTER

            'OBTENEMOS EL PIN DE SALIDA DEL UPSTREAM FILTER
            Dim pdir As PinDirection
            Dim pininfo As PinInfo
            Dim pinEnum As IEnumPins = Nothing

            Dim pinCount As Integer = 0

            Try
                USFilter.EnumPins(pinEnum)

                While pinEnum.Next(1, outputPin, IntPtr.Zero) = 0
                    outputPin(0).QueryPinInfo(pininfo)
                    If pininfo.dir = PinDirection.Output Then
                        'TENEMOS UN OUTPUT PIN, VERIFICAMOS SI COINCIDEN LOS NOMBRES

                        If USPinName = pininfo.name Then
                            'OK ESTE ES NUESTRO PIN
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pininfo.filter)
                            pinEnum = Nothing
                            bFoundPin = True
                            Exit While
                        End If
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outputPin(0))
                    DsUtils.FreePinInfo(pininfo)
                End While

                If Not bFoundPin Then
                    'NO ENCONTRAMOS EL PIN QUE BUSCABAMOS, ABANDONAMOS
                    If outputPin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(outputPin(0))
                    If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                    Return False
                End If

                'NOS PREPARAMOS PARA LA SIGUIENTE ENUMERACION
                If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                pinCount = 0
                bFoundPin = False
                DSFilter.EnumPins(pinEnum)

                While pinEnum.Next(1, inputPin, IntPtr.Zero) = 0
                    inputPin(0).QueryPinInfo(pininfo)

                    If pininfo.dir = PinDirection.Input Then
                        'TENEMOS UN PIN DE ENTRADA, VERIFICAMOS SI ES EL INDICE CORRECTO
                        If pininfo.name = DSPinName Then
                            'ESTE ES NUESTRO PIN
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pininfo.filter)
                            pinEnum = Nothing
                            bFoundPin = True
                            Exit While
                        Else
                            'SEGUIMOS CON LA ENUMERACION
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(outputPin(0))
                            DsUtils.FreePinInfo(pininfo)
                        End If
                    End If
                End While

                If Not bFoundPin Then
                    'NO ENCONTRAMOS EL PIN QUE BUSCABAMOS, ABANDONAMOS
                    If inputPin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(inputPin(0))
                    If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)
                    Return False
                End If

                If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)

                'YA TENEMOS AMBOS PINES, INTENTAMOS CONECTARLOS DIRECTAMENTE
                'PASANDO NOTHING AL MEDIA TYPE SE DEJA QUE LOS FILTROS NEGOCIEN LA CONEXION ENTRE SI
                If Me.mFG.ConnectDirect(outputPin(0), inputPin(0), Nothing) <> 0 Then
                    'NO SE PUEDEN CONECTAR AMBOS PINES (NO SE PUEDE NEGOCIAR MEDIATYPE?)
                    If bUseICIfFailure Then
                        'EL DEV DESEA UTILIZAR INTELLIGENT CONNECT EN CASO DE FALLA, LO INTENTAMOS
                        If Me.mFG.Connect(outputPin(0), inputPin(0)) = 0 Then
                            'NO SE PUDO REALIZAR LA CONEXION
                            bRet = False
                        Else
                            'SE REALIZO LA CONEXION MEDIANTE IC
                            bRet = True
                        End If
                    Else
                        'EL USUARIO NO DESEA USAR INTELLIGENT CONNECT PERO LA CONEXION DIRECTA FALLO
                        bRet = False
                    End If
                Else
                    'LA CONEXION SE REALIZO DE UNA SIN USAR INTELLIGENT CONNECT
                    bRet = True
                End If

                'LIBERAMOS LAS INTERFACES DE LOS PINES
                If inputPin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(inputPin(0))
                If outputPin(0) IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(outputPin(0))
                If pinEnum IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(pinEnum)

                Return bRet

            Catch ex As Exception
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                Return False
            End Try
        End Function

        Private Function AddCustomFilter(ByVal bIsSource As Boolean, ByVal filterCLSiD As String, ByVal filterFriendlyName As String, ByRef Filter As IBaseFilter, Optional ByVal szFileName As String = "") As Boolean
            Try
                Dim _guid As New Guid(filterCLSiD)

                'CREAMOS INSTANCIA DE OBJETO COM Y LA AGREGAMOS AL REFERENCE COUNTER
                Dim _type As Type = Type.GetTypeFromCLSID(_guid)
                Filter = DirectCast(Activator.CreateInstance(_type), IBaseFilter)

                'FINALMENTE AGREGAMOS ESA INSTANCIA AL FILTER GRAPH
                If bIsSource Then
                    If szFileName = "" Then Return False

                    DsError.ThrowExceptionForHR(Me.mFG.AddSourceFilter(szFileName, filterFriendlyName, Filter))
                Else
                    DsError.ThrowExceptionForHR(Me.mFG.AddFilter(Filter, filterFriendlyName))
                End If

                Return True
            Catch ex As Exception
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                Me.ReleaseCustomFG()
                Return False
            End Try
        End Function

        Private Function ReleaseCustomFG() As Boolean
            Try
                If Me.mSourceFilter IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(Me.mSourceFilter)
                If Me.mLAVAudioDecoder IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(Me.mLAVAudioDecoder)
                If Me.mLAVVideoDecoder IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(Me.mLAVVideoDecoder)
                If Me.mLAVSplitter IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(Me.mLAVSplitter)
                If Me.mVMR9 IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(Me.mVMR9)
                If Me.mFG IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(Me.mFG)

                Me.mSourceFilter = Nothing
                Me.mLAVAudioDecoder = Nothing
                Me.mLAVVideoDecoder = Nothing
                Me.mLAVSplitter = Nothing
                Me.mVMR9 = Nothing
                Me.mFG = Nothing
                Return True

            Catch ex As Exception
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                Return False
            End Try
        End Function

        Private Function CloseAndRelease() As Boolean
            'ESTADO INCORRECTO
            If Me.StreamState = ENUMS.ESTADO.ESTADO_CERRADO Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                Return False
            End If

            If Me.mFG Is Nothing Then
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_MODOINCORRECTO)
                Return False
            End If

            'PROCEDIMIENTO
            'RECORREMOS TODOS LOS FILTROS Y LOS ELIMINAMOS DEL FILTERGRAPH
            'ELIMINAMOS LOS OBJETOS DE LOS FILTROS Y LIBERAMOS RECURSOS
            'ELIMINAMOS EL FILTERGRAPH Y LIBERAMOS RECURSOS

            If Me.StreamState <> ENUMS.ESTADO.ESTADO_DETENIDO Then Me.Stop()

            Dim pin(0) As IPin
            Dim filter(0) As IBaseFilter
            Dim pinenum As IEnumPins
            Dim filterenum As IEnumFilters

            Me.mFG.EnumFilters(filterenum)

            While filterenum.Next(1, filter, IntPtr.Zero) = 0
                Me.mFG.RemoveFilter(filter(0))
            End While

            Marshal.ReleaseComObject(filterenum)

            Me.ReleaseCustomFG()
            Me.mAudio = Nothing
            Me.mVideo = Nothing
            Me.mControl = Nothing

            Me.StreamState = ENUMS.ESTADO.ESTADO_CERRADO
            Return True
        End Function

        Private Function CreateRectText(ByVal txt As String, ByVal ImgSize As Size, ByVal fntface As FontFamily, ByVal facecolor As Color, ByVal alphakey As Color, ByVal bBorder As Boolean, Optional ByVal bUnderline As Boolean = False, Optional ByVal bBold As Boolean = True, Optional ByVal bItalic As Boolean = False, Optional ByVal bStrikeThru As Boolean = False, Optional ByVal outlineSize As Single = 4.0, Optional ByVal fontSize As Single = 15.0F, Optional ByVal getBestFit As Boolean = True) As Image
            Dim bmp As New Bitmap(ImgSize.Width, ImgSize.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            Dim canvas As Graphics = Graphics.FromImage(bmp)
            Dim format As New StringFormat
            Dim style As Integer
            Dim tSize As Single = fontSize

            Dim g As System.Drawing.Drawing2D.GraphicsPath = New System.Drawing.Drawing2D.GraphicsPath

            Try
                'ESTABLECEMOS EL ESTILO EN BASE A LOS ARGUMENTOS
                If bBold Then style = style Or FontStyle.Bold
                If bUnderline Then style = style Or FontStyle.Underline
                If bStrikeThru Then style = style Or FontStyle.Strikeout
                If bItalic Then style = style Or FontStyle.Italic

                'CONFIGURAMOS LA CALIDAD PARA EL TEXTO
                canvas.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
                canvas.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                canvas.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                canvas.CompositingQuality = Drawing2D.CompositingQuality.HighQuality

                canvas.Clear(Color.Transparent)
                'CONFIGURAMOS ALINEACION DEL TEXTO
                format.Alignment = StringAlignment.Center
                format.LineAlignment = StringAlignment.Center
                'format.Trimming = StringTrimming.None
                format.FormatFlags = StringFormatFlags.NoClip 'Or StringFormatFlags.NoWrap
                Dim s As SizeF

                'CALCULAMOS LAS MEDIDAS DEL TEXTO Y ASI LO POSICIONAMOS RELATIVO AL TAMAÑO DE LA IMAGEN
                If getBestFit Then
                    Dim i As Single
                    While s.Width < ImgSize.Width
                        i += 1.0F
                        s = canvas.MeasureString(txt, New Font(fntface, fontSize + i)) ', New Size(ImgSize.Width, 1000), format)
                        If s.Width > ImgSize.Width Then i -= 1
                    End While
                    tSize = fontSize + i
                Else
                    s = canvas.MeasureString(txt, New Font(fntface, fontSize), ImgSize, format)
                End If

                Dim x As Single = (ImgSize.Width - s.Width) / 2
                Dim y As Single = (ImgSize.Height - s.Height)

                'PINTAMOS EL COLOR DE FONDO ESPECIFICADO EN ALPHAKEY (SE USARA POR VMR PARA LA TRANSPARENCIA)
                'canvas.FillRectangle(New SolidBrush(alphakey), 0, 0, ImgSize.Width, ImgSize.Height)

                'AGREGAMOS EL TEXTO CON EL ESTILO ESPECIFICADO
                g.AddString(txt, fntface, style, tSize, New Rectangle(x, y, s.Width, s.Height), format)

                'CONFIGURAMOS PARAMETROS PARA EL OUTLINE, EN ESTE CASO SERA NEGRO CON BORDES REDONDEADOS
                Dim outlinePen As Pen

                If bBorder Then
                    outlinePen = New Pen(Brushes.Black, outlineSize)
                Else
                    outlinePen = New Pen(facecolor, 0.100000001F)
                End If

                outlinePen.LineJoin = Drawing2D.LineJoin.Round

                'ESCRIBIMOS EL OUTLINE SEGUIDO DEL TEXTO
                canvas.DrawPath(outlinePen, g)
                If facecolor = Color.Black Then facecolor = Color.DimGray 'NO PERMITIMOS COLOR NEGRO PORQUE EXPLOTA
                canvas.FillPath(New SolidBrush(facecolor), g)

                'canvas.DrawString(txt, fntface, New SolidBrush(facecolor), New RectangleF(x, y, s.Width, s.Height), format))

                canvas.Dispose()
                format.Dispose()
                g.Dispose()
                outlinePen.Dispose()

                Return bmp
            Catch ex As Exception
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DESCONOCIDO)
                Return Nothing
            End Try
        End Function

        Private Function RenderIntelligentConnect(ByVal szFileName As String) As Boolean
            Try
                If Me.mFG Is Nothing Then Return False

                'AGREGAMOS VMR9
                If Me.AddCustomFilter(False, "51B4ABF3-748F-4E3B-A276-C828330E926A", "Video Mixing Renderer 9", Me.mVMR9) = False Then Return False

                'CONFIGURAMOS EL RENDERER
                Me.ConfigureRenderer()

                DsError.ThrowExceptionForHR(Me.mFG.RenderFile(szFileName, vbNullString))
                Me.mIsUsingIC = True
                RaiseEvent UsingIC()
                Return True
            Catch ex As Exception
                Me.mIsUsingIC = False
                RaiseEvent Error(ENUMS.ERROR_LIST.ERL_DIRECTSHOW)
                Return False
            End Try
        End Function
#End Region

#Region "IDisposable Support"
        Private disposedValue As Boolean ' Para detectar llamadas redundantes

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: eliminar estado administrado (objetos administrados).
                    Me.ReleaseDirect3D()
                End If

                ' TODO: liberar recursos no administrados (objetos no administrados) e invalidar Finalize() below.
                ' TODO: Establecer campos grandes como Null.
                Me.Close()
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: invalidar Finalize() sólo si la instrucción Dispose(ByVal disposing As Boolean) anterior tiene código para liberar recursos no administrados.
        Protected Overrides Sub Finalize()
            ' No cambie este código. Ponga el código de limpieza en la instrucción Dispose(ByVal disposing As Boolean) anterior.
            Dispose(False)
            MyBase.Finalize()
        End Sub

        ' Visual Basic agregó este código para implementar correctamente el modelo descartable.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' No cambie este código. Coloque el código de limpieza en Dispose (ByVal que se dispone como Boolean).
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class
End Namespace
