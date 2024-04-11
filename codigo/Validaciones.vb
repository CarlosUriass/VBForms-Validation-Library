Public Class Validaciones

    Public Shared Function validacion_vacio(ByVal texto As String) As Boolean
        '-------------------------------------------------------------------------------
        ' Función: validacion_vacio
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si un texto dado está vacío o no. Si el texto está vacío,
        '   muestra un mensaje de advertencia y devuelve True; de lo contrario, devuelve False.
        '
        ' Parámetros:
        '   texto (String): El texto que se va a verificar si está vacío o no.
        '
        ' Devuelve:
        '   Boolean: True si el texto está vacío, False si no lo está.
        '-------------------------------------------------------------------------------

        ' Verifica si el texto está vacío
        If texto = "" Then
            ' Si el texto está vacío, muestra un mensaje de advertencia
            MsgBox("No puede haber campos vacíos", MsgBoxStyle.OkOnly)
            Return True
        Else
            ' Si el texto no está vacío, devuelve False
            Return False
        End If
    End Function



    Public Shared Function validarLongitud(ByVal texto As String, ByVal longitud As Integer) As Boolean

        '-------------------------------------------------------------------------------
        ' Función: validarLongitud
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si un texto dado tiene una longitud especificada. Si la longitud
        '   del texto no coincide con la longitud especificada, muestra un mensaje de advertencia y
        '   devuelve True; de lo contrario, devuelve False.
        '
        ' Parámetros:
        '   texto (String): El texto cuya longitud se va a verificar.
        '   longitud (Integer): La longitud requerida del texto que se va a verificar.
        '
        ' Devuelve:
        '   Boolean: True si la longitud del texto no es igual a la longitud especificada,
        '            False si lo es.
        '-------------------------------------------------------------------------------


        ' Verifica si la longitud del texto es diferente de la longitud especificada
        If Len(texto) <> longitud Then
            ' Si la longitud del texto no es exactamente la longitud especificada, devuelve True y muestra un mensaje de advertencia
            MsgBox("El campo debe tener exactamente " & longitud & " caracteres de longitud", MsgBoxStyle.OkOnly)
            Return True
        Else
            ' Si la longitud coincide con la longitud especificada, devuelve False
            Return False
        End If
    End Function

    '-------------------------------------------------------------------------------
    ' Función: soloNumeros
    '-------------------------------------------------------------------------------
    ' Descripción:
    '   Esta función verifica si un texto dado contiene solo caracteres numéricos. 
    '   Si encuentra algún carácter que no sea un número, muestra un mensaje de advertencia
    '   y devuelve True; de lo contrario, devuelve False.
    '
    ' Parámetros:
    '   texto (String): El texto que se va a verificar para asegurar que contenga solo números.
    '
    ' Devuelve:
    '   Boolean: True si el texto contiene caracteres que no son números, False si todos los caracteres son números.
    '-------------------------------------------------------------------------------

    Public Shared Function soloNumeros(ByVal texto As String) As Boolean
        ' Validar que solo sean números
        For i As Integer = 1 To Len(texto)
            If Not Char.IsDigit(texto(i - 1)) Then
                ' Si encuentra un carácter que no es un número, muestra un mensaje de advertencia y devuelve True
                MsgBox("El campo solo debe contener números", MsgBoxStyle.OkOnly)
                Return True
                Exit For ' Salir del bucle luego de encontrar el primer carácter no numérico
            End If
        Next
        Return False
    End Function

    Public Shared Function caracteresEspeciales(ByVal texto As String) As Boolean
        '-------------------------------------------------------------------------------
        ' Función: caracteresEspeciales
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si un texto contiene caracteres especiales. Si encuentra
        '   algún carácter especial, muestra un mensaje de advertencia y devuelve True; de lo
        '   contrario, devuelve False.
        '
        ' Parámetros:
        '   texto (String): El texto que se va a verificar en busca de caracteres especiales.
        '
        ' Devuelve:
        '   Boolean: True si el texto contiene caracteres especiales, False si no los contiene.
        '
        ' Caracteres Especiales en la Tabla ASCII:
        '   32  (space)
        '   33  !
        '   34  "
        '   35  #
        '   36  $
        '   37  %
        '   38  &
        '   39  '
        '   40  (
        '   41  )
        '   42  *
        '   43  +
        '   44  ,
        '   47  /
        '   58  :
        '   59  ;
        '   60  <
        '   61  =
        '   62  >
        '   63  ?
        '   64  @
        '   91  [
        '   92  \
        '   93  ]
        '   94  ^
        '   123 {
        '   124 |
        '   125 }
        '   126 ~

        'Nota: El 45 (-), 46(.) y el 95(_) fueron removidos para tener compatibilidad con la validación de correo electrónico
        '-------------------------------------------------------------------------------

        For i As Integer = 1 To Len(texto)
            Select Case Asc(texto(i - 1))
                Case 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 47, 58, 59, 60, 61, 62, 63, 64, 91, 92, 93, 94, 123, 124, 125, 126
                    MsgBox("El campo no puede contener caracteres especiales", MsgBoxStyle.OkOnly)
                    Return True
                    Exit For
            End Select
        Next

        Return False
    End Function


    Public Shared Function validarDinero(ByVal monto As String) As Boolean
        '-------------------------------------------------------------------------------
        ' Función: validarDinero
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si una cadena dada representa un monto de dinero válido.
        '   Verifica si la cadena contiene solo números y opcionalmente un punto decimal,
        '   lo que indica la separación entre la parte entera y la parte decimal del monto.
        '
        ' Parámetros:
        '   monto (String): La cadena que representa el monto de dinero que se va a validar.
        '
        ' Devuelve:
        '   Boolean: True si la cadena no cumple con los criterios de validación, False si
        '            cumple con los criterios y es considerada un monto de dinero válido.
        '
        ' Dependencias:
        '   Esta función hace uso de la función soloNumeros previamente definida para verificar
        '   si una cadena contiene solo caracteres numéricos.
        '-------------------------------------------------------------------------------
        Dim punto_posicion As Byte

        ' Buscar la posición del punto decimal en la cadena
        For i As Integer = 1 To Len(monto)
            If monto(i - 1) = "." Then
                punto_posicion = i
                Exit For
            End If
        Next

        ' Validar que solo sean números antes del punto decimal
        If soloNumeros(Mid(monto, 1, punto_posicion - 1)) Then
            Return True
        End If

        ' Validar que solo sean números después del punto decimal (si existe)
        If punto_posicion > 0 AndAlso soloNumeros(Mid(monto, punto_posicion + 1)) Then
            Return True
        End If

        ' Si todas las validaciones pasan, entonces el monto de dinero es válido
        Return False
    End Function



    Public Shared Function mayorquecero(ByVal numero As Integer) As Boolean
        '-------------------------------------------------------------------------------
        ' Función: mayorquecero
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si un número entero es mayor que cero. Si el número es
        '   menor que cero, muestra un mensaje de advertencia y devuelve True; de lo contrario,
        '   devuelve False.
        '
        ' Parámetros:
        '   numero (Integer): El número entero que se va a verificar si es mayor que cero.
        '
        ' Devuelve:
        '   Boolean: True si el número es menor que cero, False si es mayor o igual a cero.
        '-------------------------------------------------------------------------------

        If numero < 0 Then
            MsgBox("El campo no puede ser menor que 0", MsgBoxStyle.OkOnly)
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function longitudMaxima(ByVal texto As String, maximo As Integer) As Boolean

        '-------------------------------------------------------------------------------
        ' Función: longitudMaxima
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si la longitud de un texto dado excede un límite máximo
        '   especificado. Si la longitud del texto es mayor que el límite máximo, muestra
        '   un mensaje de advertencia y devuelve True; de lo contrario, devuelve False.
        '
        ' Parámetros:
        '   texto (String): El texto cuya longitud se va a verificar.
        '   maximo (Integer): La longitud máxima permitida para el texto.
        '
        ' Devuelve:
        '   Boolean: True si la longitud del texto excede el límite máximo, False si no lo hace.
        '-------------------------------------------------------------------------------

        If Len(texto) > maximo Then
            MsgBox("El campo excede los caracteres permitidos", MsgBoxStyle.OkOnly)
            Return True
        End If

        Return False


    End Function

    Public Shared Function longitudMinima(ByVal texto As String, minimo As Integer) As Boolean

        '-------------------------------------------------------------------------------
        ' Función: longitudMinima
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si la longitud de un texto dado es menor que un límite
        '   mínimo especificado. Si la longitud del texto es menor que el límite mínimo,
        '   muestra un mensaje de advertencia y devuelve True; de lo contrario, devuelve False.
        '
        ' Parámetros:
        '   texto (String): El texto cuya longitud se va a verificar.
        '   minimo (Integer): La longitud mínima requerida para el texto.
        '
        ' Devuelve:
        '   Boolean: True si la longitud del texto es menor que el límite mínimo, False si no lo es.
        '-------------------------------------------------------------------------------


        If Len(texto) < minimo Then
            MsgBox("El campo debe tener al menos " & minimo & " caracteres", MsgBoxStyle.OkOnly)
            Return True
        End If


        Return False
    End Function


    Public Shared Function numerotelefonico(ByVal numero As String) As Boolean
        '-------------------------------------------------------------------------------
        ' Función: numerotelefonico
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si un número telefónico dado cumple con los criterios
        '   de validación establecidos. Primero, verifica si el número contiene solo
        '   caracteres numéricos. Luego, verifica si el número tiene una longitud de
        '   10 dígitos, que es la longitud estándar para números telefónicos en México.
        '
        ' Parámetros:
        '   numero (String): El número telefónico que se va a validar.
        '
        ' Devuelve:
        '   Boolean: True si el número telefónico cumple con los criterios de validación,
        '            False si no los cumple.
        '
        ' Dependencias:
        '   Esta función utiliza las siguientes funciones de validación definidas previamente:
        '   - soloNumeros: Verifica si una cadena contiene solo caracteres numéricos.
        '   - validarLongitud: Verifica si una cadena tiene una longitud específica.
        '-------------------------------------------------------------------------------

        'Primero validar que sean puros números

        If soloNumeros(numero) = True Then
            Return True
        End If

        'Validar longitud

        If validarLongitud(numero, 10) = True Then
            Return True
        End If


        Return False

    End Function

    Public Shared Function correoelectronico(ByVal email As String) As Boolean

        '-------------------------------------------------------------------------------
        ' Función: correoelectronico
        '-------------------------------------------------------------------------------
        ' Descripción:
        '   Esta función verifica si una cadena dada cumple con los criterios para ser
        '   considerada como una dirección de correo electrónico válida. Verifica si la
        '   cadena tiene un formato adecuado, incluyendo la presencia de una única arroba,
        '   longitud del nombre de usuario y del dominio dentro de los límites permitidos,
        '   ausencia de caracteres especiales en el nombre de usuario, y si el dominio
        '   tiene una extensión de dominio (TLD) válida (.com, .gob, .org, .net).
        '
        ' Parámetros:
        '   email (String): La cadena que representa la dirección de correo electrónico
        '                   que se va a validar.
        '
        ' Devuelve:
        '   Boolean: True si la cadena no cumple con los criterios de validación, False si
        '            cumple con los criterios y es considerada una dirección de correo
        '            electrónico válida.
        '
        ' Dependencias:
        '   Esta función hace uso de las siguientes funciones de validación previamente
        '   definidas:
        '   - longitudMaxima: Verifica si una cadena excede una longitud máxima especificada.
        '   - longitudMinima: Verifica si una cadena tiene una longitud mínima especificada.
        '   - caracteresEspeciales: Verifica si una cadena contiene caracteres especiales.
        '-------------------------------------------------------------------------------


        Dim arroba_posc As Integer
        Dim contador_arroba As Integer
        contador_arroba = 0

        ' Contar la cantidad de arrobas en el correo electrónico
        For i As Integer = 1 To Len(email)
            If email(i - 1) = "@" Then
                arroba_posc = i
                contador_arroba += 1
            End If
        Next

        If contador_arroba > 1 Then
            MsgBox("Solo debe existir una arroba en el email", MsgBoxStyle.OkOnly)
            Return True
        End If

        If contador_arroba = 0 Then
            MsgBox("El correo debe contener al menos un arroba", MsgBoxStyle.OkOnly)
            Return True
        End If

        'Generar nombre de usuario
        Dim nombre_usuario As String
        nombre_usuario = Microsoft.VisualBasic.Left(email, arroba_posc - 1)

        'Validar si el nombre de usuario excede los 64 caracteres

        If longitudMaxima(nombre_usuario, 64) = True Then
            Return True
        End If

        'Validar que el nombre de usuario sea mayor que 1

        If longitudMinima(nombre_usuario, 2) = True Then
            Return True
        End If

        'Generar dominio
        Dim dominio As String
        dominio = Mid(email, arroba_posc + 1)

        'Validar que el dominio no exceda los 255 caracteres

        If longitudMaxima(dominio, 64) = True Then
            Return True
        End If

        'Validar que el nombre de usuario no tenga caracteres especiales

        If caracteresEspeciales(nombre_usuario) = True Then
            Return True
        End If

        ' Validar que el dominio tenga una extensión de dominio (TLD) válida
        If Not (dominio.EndsWith(".com") Or dominio.EndsWith(".gob") Or dominio.EndsWith(".org") Or dominio.EndsWith(".net")) Then
            MsgBox("El dominio del correo electrónico no es válido.", MsgBoxStyle.OkOnly)
            Return True
        End If


        Return False

    End Function

End Class
