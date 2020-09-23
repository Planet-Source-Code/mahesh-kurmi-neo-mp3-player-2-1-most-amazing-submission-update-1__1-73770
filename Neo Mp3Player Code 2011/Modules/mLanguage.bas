Attribute VB_Name = "mLanguage"
Option Explicit

'Public arryLanguage() As String

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  Implementing later on                                                                               |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Load_Language_Spanish()
    With frmPopUp
        .lstLanguage.Clear
        .lstLanguage.AddItem "Language"
        ' MENU
        .lstLanguage.AddItem " Nueva Busqueda"  ' 1
        .lstLanguage.AddItem " Caratula"    ' 2
        .lstLanguage.AddItem "  Cambiar ListaRep / Caratula"    ' 3
        .lstLanguage.AddItem "  Colocar caratula como Wallpaper"    ' 4
        .lstLanguage.AddItem "  Maximizar Caratula"    '  5
        .lstLanguage.AddItem " Exploradores"    '  6
        .lstLanguage.AddItem "  Explorar archivos"    ' 7
        .lstLanguage.AddItem "   Explorador de Albums"    ' 8
        .lstLanguage.AddItem "   Editar Track(s) Tag"    '  9
        .lstLanguage.AddItem "   Karaoke"    ' 10
        .lstLanguage.AddItem " Visualización Studio"    ' 11
        .lstLanguage.AddItem "   Configurar Visualización"    ' 12
        .lstLanguage.AddItem "   Mostrar Visualización"    ' 13
        .lstLanguage.AddItem " Controles de Reproducción"    ' 14
        .lstLanguage.AddItem "   Volumen"    ' 15
        .lstLanguage.AddItem "+     Subir Volumen"    ' 16
        .lstLanguage.AddItem "-     Bajar Volumen"    ' 17
        .lstLanguage.AddItem "Z   Track Anterior"    ' 18
        .lstLanguage.AddItem "X   Reproducir"    ' 19
        .lstLanguage.AddItem "C   Pausar"    ' 20
        .lstLanguage.AddItem "V   Detener"    ' 21
        .lstLanguage.AddItem "B   Siguiente Track"    ' 22
        .lstLanguage.AddItem "<   Anterior Album / Folder"    ' 23
        .lstLanguage.AddItem ">   Siguiente Album / Folder"    ' 24
        .lstLanguage.AddItem "I   Intro 10 seg."    ' 25
        .lstLanguage.AddItem "R   Repetir Track"    ' 26
        .lstLanguage.AddItem "S   Silencio"    ' 27
        .lstLanguage.AddItem "   Orden aleatorio"    ' 28
        .lstLanguage.AddItem "Q     Actual Album / Folder"    ' 29
        .lstLanguage.AddItem "W     Todos los Albums"    ' 30
        .lstLanguage.AddItem "A   Atras 5 seg."    ' 31
        .lstLanguage.AddItem "D   Adelante 5 Seg."    ' 32
        .lstLanguage.AddItem " Opciones"    ' 33
        .lstLanguage.AddItem " Skins"    ' 34
        .lstLanguage.AddItem "   << Explorador de Skins >>"    ' 35
        .lstLanguage.AddItem " Transparencia"    ' 36
        .lstLanguage.AddItem "   Personalizar"    ' 37
        .lstLanguage.AddItem " Acerca de"    ' 38
        .lstLanguage.AddItem " Salir"    ' 39
        ' ACERCA
        .lstLanguage.AddItem " Acerca de MaheshMp3 Player"    ' 40
        ' CARATULA
        .lstLanguage.AddItem "Caratula actual"
        ' EXPLORADOR DE ALBUMS
        .lstLanguage.AddItem "Explorador de Albums"
        .lstLanguage.AddItem "  Explorar archivos"
        .lstLanguage.AddItem "  Editar Tags"
        .lstLanguage.AddItem "  Reproducir"
        ' KARAOKE
        .lstLanguage.AddItem "Karaoke"
        .lstLanguage.AddItem "  [ Letras no Encontradas ]"
        ' MAIN
        .lstLanguage.AddItem "    Menu"
        .lstLanguage.AddItem "    Minimizar"
        .lstLanguage.AddItem "    Change Mode"    ' 50
        .lstLanguage.AddItem "    Salir"
        .lstLanguage.AddItem "  Selecciona el directorio a buscar."
        .lstLanguage.AddItem "  Sin Visualización"
        .lstLanguage.AddItem "  Analizador de Espectro"
        .lstLanguage.AddItem "  Osiloscopio"
        .lstLanguage.AddItem "  Editar Album Tags"
        .lstLanguage.AddItem "  Explorar Exp. de Albums"
        .lstLanguage.AddItem "  Explorar Explorer.exe"
        .lstLanguage.AddItem "  Reproducir"
        ' TAGS
        .lstLanguage.AddItem "Editor de Tags + Información MPEG"    ' 60
        .lstLanguage.AddItem "  Multiples tracks estan seleccionados, Selecciona los checkboxs para aplicar los cambios  a TODOS los archivos seleccionados."
        .lstLanguage.AddItem "  Seleccionar Todo"
        .lstLanguage.AddItem "  Tags"
        .lstLanguage.AddItem "  Karaoke"
        .lstLanguage.AddItem "  Agregar"
        .lstLanguage.AddItem "  Deshacer"
        .lstLanguage.AddItem "  Aceptar"
        .lstLanguage.AddItem "  Cancelar"
        .lstLanguage.AddItem "  Aplicar"
        ' OPCIONES
        .lstLanguage.AddItem "Opciones"    ' 70
        .lstLanguage.AddItem "  Aceptar"    ' 71
        .lstLanguage.AddItem "  Cancelar"    ' 72
        .lstLanguage.AddItem "  Aplicar"    ' 73
        .lstLanguage.AddItem "Aplicación"    ' 74
        .lstLanguage.AddItem "  Trayectoria Configuración"    ' 75
        .lstLanguage.AddItem "    Trayectoria de skins y configuración:"    ' 76
        .lstLanguage.AddItem "    Explorar..."    ' 77
        .lstLanguage.AddItem "    Nota: Algunas opciones requieren que se reinicie la aplicación."
        .lstLanguage.AddItem "    Memoria Libre (Fisica):"    ' 79
        .lstLanguage.AddItem "  Aplicación"    ' 80
        .lstLanguage.AddItem "    Lenguaje:"    ' 81
        .lstLanguage.AddItem "    Siempre arriba."    ' 82
        .lstLanguage.AddItem "    Mostrar Splash Screen."    ' 83
        .lstLanguage.AddItem "    Permitir multiples instancias."    ' 84
        .lstLanguage.AddItem "    Habilitar menu en drives y directorios."    ' 85
        .lstLanguage.AddItem "    Mostrar MaheshMp3 Player en:"    ' 86
        .lstLanguage.AddItem "    Barra de tareas."    ' 87
        .lstLanguage.AddItem "    Bandeja de sistema."    ' 88
        .lstLanguage.AddItem "  Transparencia"    ' 89
        .lstLanguage.AddItem "    Transparencia(Solo win 2000 o sup.)"    ' 90
        .lstLanguage.AddItem "Skins"    ' 91
        .lstLanguage.AddItem "  Skin actual:"    ' 92
        .lstLanguage.AddItem "  Información:"    ' 93
        .lstLanguage.AddItem "  Cargar región desde archivo."    ' 94
        .lstLanguage.AddItem "Wallpaper"    ' 95
        .lstLanguage.AddItem "  Opciones de fondo de escritorio."    ' 96
        .lstLanguage.AddItem "  No alterar."    ' 97
        .lstLanguage.AddItem "  Ajustar."    ' 98
        .lstLanguage.AddItem "  Centrar."    ' 99
        .lstLanguage.AddItem "  Mosaico."    ' 100
        .lstLanguage.AddItem "  Proporcional."    ' 101
        .lstLanguage.AddItem "Play List"    ' 102
        .lstLanguage.AddItem "  Formato Lista de Reproducción."    ' 103
        .lstLanguage.AddItem "  Formato de texto reproduciendo."    ' 104
        .lstLanguage.AddItem "  Tipo de Scroll:"    ' 105
        .lstLanguage.AddItem "    Rotar."    ' 106
        .lstLanguage.AddItem "    Zig Zag."    ' 107
        .lstLanguage.AddItem "  Velocidad del Scroll:"    ' 108
        .lstLanguage.AddItem "Reproductor"    ' 109
        .lstLanguage.AddItem "  Reproducir archivos:"    ' 110
        .lstLanguage.AddItem "  Mostrar icono en bandeja de sistema:"    ' 111
        .lstLanguage.AddItem "    Anterior Track."    ' 112
        .lstLanguage.AddItem "    Reproducir."    ' 113
        .lstLanguage.AddItem "    Pausar."    ' 114
        .lstLanguage.AddItem "    Detener."    ' 115
        .lstLanguage.AddItem "    Siguente Track."    ' 116
        .lstLanguage.AddItem "  Crossfade entre tracks (ms):"    ' 117
        .lstLanguage.AddItem "  Crossfade en Detener (ms):"    ' 118
        .lstLanguage.AddItem "  Reproducir al inicio."    ' 119
        .lstLanguage.AddItem "Efectos FX"    ' 120
        .lstLanguage.AddItem "  coro"    ' 121
        .lstLanguage.AddItem "    Habilitar coro."    ' 122
        .lstLanguage.AddItem "      Mezcla:"    ' 123
        .lstLanguage.AddItem "      Profundidad:"    ' 124
        .lstLanguage.AddItem "      Retroacción:"    ' 125
        .lstLanguage.AddItem "      Frecuencía:"    ' 126
        .lstLanguage.AddItem "      Forma de onda:"    ' 127
        .lstLanguage.AddItem "      Retrazo:"    ' 128
        .lstLanguage.AddItem "      Fase:"    ' 129
        .lstLanguage.AddItem "  compresor"    ' 130
        .lstLanguage.AddItem "    Habilitar compresor."    ' 131
        .lstLanguage.AddItem "      Incremento:"    ' 132
        .lstLanguage.AddItem "      Ataque:"
        .lstLanguage.AddItem "      Edición:"
        .lstLanguage.AddItem "      Umbral:"
        .lstLanguage.AddItem "      Proporción:"
        .lstLanguage.AddItem "      Preretrazo:"
        .lstLanguage.AddItem "  Distorción"
        .lstLanguage.AddItem "    Habilitar Distorción:"
        .lstLanguage.AddItem "      Incremento:"    ' 140
        .lstLanguage.AddItem "      Bordes:"
        .lstLanguage.AddItem "      frecuencia Central:"
        .lstLanguage.AddItem "      Ancho frecuencia:"
        .lstLanguage.AddItem "      Atenuación:"
        .lstLanguage.AddItem "  eco"    '145
        .lstLanguage.AddItem "    Habilitar eco."
        .lstLanguage.AddItem "      Mezcla:"
        .lstLanguage.AddItem "      Retroaccion:"
        .lstLanguage.AddItem "      Atraso izquierda:"
        .lstLanguage.AddItem "      Atraso derecha:"
        .lstLanguage.AddItem "      Atraso Central:"
        .lstLanguage.AddItem "  Flanger"    ' 152
        .lstLanguage.AddItem "    Habilitar Flanger."
        .lstLanguage.AddItem "      Mezcla:"
        .lstLanguage.AddItem "      Profundidad:"
        .lstLanguage.AddItem "      Retroaccion:"
        .lstLanguage.AddItem "      Frecuencia:"
        .lstLanguage.AddItem "      Forma de Onda:"
        .lstLanguage.AddItem "      Retrazo:"
        .lstLanguage.AddItem "      Fase:"    '160
        .lstLanguage.AddItem "  Gargarizar"    ' 161
        .lstLanguage.AddItem "    Habilitar gargarizar."
        .lstLanguage.AddItem "      Hz:"
        .lstLanguage.AddItem "      Forma de Onda:"
        .lstLanguage.AddItem "  I3DL2 Reverberación"    ' 165
        .lstLanguage.AddItem "    Habilitar I3D nivel 2 Reverberación."
        .lstLanguage.AddItem "      Cuarto:"
        .lstLanguage.AddItem "      Cuarto HF:"
        .lstLanguage.AddItem "      Factor giratorio:"    '169
        .lstLanguage.AddItem "      Tiempo decadencia:"
        .lstLanguage.AddItem "      Prop. dec. HF:"
        .lstLanguage.AddItem "      Reflecciones:"
        .lstLanguage.AddItem "      Atraso Refleccción:"
        .lstLanguage.AddItem "      Reverberación:"
        .lstLanguage.AddItem "      Atraso de Rev.:"
        .lstLanguage.AddItem "      Difusión:"
        .lstLanguage.AddItem "      Densidad:"
        .lstLanguage.AddItem "      HF Referencia:"
        .lstLanguage.AddItem "  Reverberación"    '179
        .lstLanguage.AddItem "    Habilitar Reverberación de ondas."
        .lstLanguage.AddItem "      Incremento:"
        .lstLanguage.AddItem "      Mezcla Reverberación:"
        .lstLanguage.AddItem "      Tiempo de Rev.:"
        .lstLanguage.AddItem "      HF Proporción:"
        .lstLanguage.AddItem "  Valores por default"    ' 185
        .lstLanguage.AddItem "  Desabilitar todos"
        .lstLanguage.AddItem "Equalizador"
        .lstLanguage.AddItem "  Habilitar EQ."
        .lstLanguage.AddItem "  Presentes:"    '190
        .lstLanguage.AddItem "  Borrar EQ"
        .lstLanguage.AddItem "  Guardar EQ"
        .lstLanguage.AddItem "  Nombre del Equalizador:"
        .lstLanguage.AddItem "  Borrar equalizador:"
        .lstLanguage.AddItem "Visualización"    '195
        .lstLanguage.AddItem "  Visualizaciones:"
        .lstLanguage.AddItem "  Presentes:"
        .lstLanguage.AddItem "  Nuevos:"
        .lstLanguage.AddItem "  Tipo Fondo:"    '198
        .lstLanguage.AddItem "  Peaks:"    '200
        .lstLanguage.AddItem "  Barras:"
        .lstLanguage.AddItem "  Archivo Imagen:"
        .lstLanguage.AddItem "  Escala:"
        .lstLanguage.AddItem "  Color Barras:"
        .lstLanguage.AddItem "  Num. Barras:"    '205
        .lstLanguage.AddItem "  Espacio:"
        .lstLanguage.AddItem "  Reflejo:"
        .lstLanguage.AddItem "  Color Peak:"
        .lstLanguage.AddItem "  Alto Peak:"
        .lstLanguage.AddItem "  Gravedad Peak:"    '210
        .lstLanguage.AddItem "  Gradiente:"
        .lstLanguage.AddItem "  Color Fondo:"
        .lstLanguage.AddItem "  Color Linea:"
        .lstLanguage.AddItem "  Num. Lineas"
        .lstLanguage.AddItem "  Alineacion:"    '215
        .lstLanguage.AddItem "  Guardar"
        .lstLanguage.AddItem "  Guardar como"
        .lstLanguage.AddItem "  Borrar"
        .lstLanguage.AddItem "  Mostrar"    '219
        .lstLanguage.AddItem "  Borrar Visualización:"
        .lstLanguage.AddItem "  Nombre de Visualización:"
        .lstLanguage.AddItem "  Anterior Visualizacion"
        .lstLanguage.AddItem "  Siguiente Visualizacion"
        .lstLanguage.AddItem "  Configurar ..."
        .lstLanguage.AddItem "  Salir"
        .lstLanguage.AddItem "  Guardar Config."
        .lstLanguage.AddItem "  Física:"
        .lstLanguage.AddItem "  Virtual:"
        .lstLanguage.AddItem "  Archivo:"
        .lstLanguage.AddItem " Buscar Archivos de sonido."
        .lstLanguage.AddItem " Buscar en:"
        .lstLanguage.AddItem " Explorar..."
        .lstLanguage.AddItem " Comenzar a Buscar"
        .lstLanguage.AddItem " Detener Busqueda"
        .lstLanguage.AddItem " Agregar Nuevos Archivos."

    End With
End Sub


Public Sub Load_Language(strLang As String)
    On Error Resume Next
    Dim Linenr As Integer
    Dim InputData
    Dim strRuta As String, strTemp As String
    With frmPopUp

        strRuta = tAppConfig.AppConfig & "Language\" & strLang & ".lng"
        Load_Language_Spanish
        If Dir(strRuta) <> "" Then
            Open strRuta For Input As #2

            Linenr = 0
            Do While Not EOF(2)
                Line Input #2, InputData

                If Linenr > 234 Then
                    Exit Do
                End If
                If Trim(InputData) <> "" And Linenr > 0 Then

                    If Linenr > 15 And Linenr < 33 And Linenr <> 28 Then
                        strTemp = Left(LineLanguage(Linenr), 1)
                        strTemp = Trim(strTemp) & "" & InputData
                        .lstLanguage.list(Linenr) = Trim(strTemp)
                    Else
                        .lstLanguage.list(Linenr) = Trim(InputData)
                    End If
                End If

                Linenr = Linenr + 1
            Loop
            Close #2
        End If
        ' MENU
        '.mnuNuevaBusqueda.Caption = LineLanguage(1)
        ' .mnuCFront.Caption = LineLanguage(2)
        '.mnuCambiarListaCaratula.Caption = LineLanguage(3)
        'frmMain.Button(10).ToolTipText = Trim(LineLanguage(3))
        '.mnuWallpapper.Caption = LineLanguage(4)
        '.mnuMCaratula.Caption = LineLanguage(5)
        '.mnuBrowsers.Caption = LineLanguage(6)







        '//change language at systray icons
    End With
End Sub






'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Sub Load_Language_Tags()
    Exit Sub
    With frmTags
        .Caption = LineLanguage(60)
        ' .cmdOk.Caption = LineLanguage(67)
        '.cmdCancel.Caption = LineLanguage(68)
        .vkBtn(2).Caption = LineLanguage(69)
        ' .cmdSelAll.Caption = LineLanguage(62)
        ' .TabStrip.Tabs(1).Caption = LineLanguage(63)
        ' .TabStrip.Tabs(2).Caption = LineLanguage(64)
        ' .cmdAdd.Caption = LineLanguage(65)
        ' .cmdUndo.Caption = LineLanguage(66)
    End With
End Sub

Public Function LineLanguage(Number As Integer) As String
    On Error Resume Next

    If frmPopUp.lstLanguage.ListCount = 0 Then Exit Function
    If Number > frmPopUp.lstLanguage.ListCount - 1 Then Exit Function
    LineLanguage = Trim(frmPopUp.lstLanguage.list(Number))

End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

